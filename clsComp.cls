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
Private sWrkbkBaseName      As String       ' Common Component host Workbook's base name
Private sWrkbkFullName      As String       ' Common Component host Workbook's full name
Private sCompName           As String       ' The class module's global component name
Private wb                  As Workbook     ' The class' Workbook
Private vbc                 As VBComponent  ' VBComponent managed by this Class Module
Private dtUpdatedAsOf       As Date
Private flExpFile           As FILE
Private foExport            As Folder
Private sPathCommon         As String
Private sExpFileFullName    As String       ' Full filespec for the Export File created by the method BackUpCode
Private sExpPath            As String
Private lCompMaxLen         As Long
Private lKindOfComp         As enKindOfComp
Private lKindOfCodeChange   As enKindOfCodeChange
Private sExtension          As String

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
    
    On Error GoTo eh
    sPathCommon = mMe.VBProjectsDevRoot
    Set dctRawComponents = mRaw.Components
    Set dctRawHosts = mHost.Hosts

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Property Get CompName() As String:                       CompName = sCompName:                   End Property

Public Property Let CompName(ByVal s As String)
' ---------------------------------------------
' Provides the properties:
' - CompName (sCompName)
' - VBComp (vbc)
' - Extension (sExtension),
' - ExpFileFullName (sExpFileFullName)
' ---------------------------------------------

    sCompName = s
    If Not wb Is Nothing Then
        Set vbc = wb.VBProject.VBComponents(sCompName)
        Select Case vbc.Type
            Case vbext_ct_StdModule:    sExtension = ".bas"
            Case vbext_ct_ClassModule:  sExtension = ".cls"
            Case vbext_ct_MSForm:       sExtension = ".frm"
            Case vbext_ct_Document:     sExtension = ".cls"
        End Select
        sExpFileFullName = sExpPath & "\" & sCompName & sExtension
    End If

End Property

Public Property Get ExpFile() As FILE
    
    Dim fso As New FileSystemObject
    
    If flExpFile Is Nothing Then
        If fso.FileExists(sExpFileFullName) Then
            Set flExpFile = GetFile(sExpFileFullName)
        Else
            Me.BackUpCode
            Set flExpFile = GetFile(sExpFileFullName)
        End If
    End If
    
    Set ExpFile = flExpFile
    Set fso = Nothing
    
End Property

Public Property Let ExpFile(ByVal fl As FILE):               Set flExpFile = fl:                      End Property

Private Property Get ExpFileExtension() As String
    Select Case vbc.Type
        Case vbext_ct_StdModule:    ExpFileExtension = ".bas"
        Case vbext_ct_ClassModule:  ExpFileExtension = ".cls"
        Case vbext_ct_MSForm:       ExpFileExtension = ".frm"
        Case vbext_ct_Document:     ExpFileExtension = ".cls"
    End Select
End Property

Public Property Get ExpFileFullName() As String:            ExpFileFullName = sExpFileFullName:         End Property

Public Property Get ExpFolder() As Folder:                  Set ExpFolder = foExport:                   End Property

Public Property Get Extension() As String:                  Extension = sExtension:                     End Property

Public Property Get KindOfCodeChange( _
           Optional ignore_empty_lines As Boolean = False) As enKindOfCodeChange
' ------------------------------------------------------------------------------
' Returns the kind of code change.
' - The "Raw" has changed when the Component is a used Raw and its Export File
'   differs from the Export File of the Raw Component hosted in another Workbook
' - The "Used Raw Component" or the "Neither Hosted Nor Used Raw Component" has
'   changed when the code of the VBComponent differs from its Export File.
' - Both, the "Raw Hosted in another Workbook" and the "Used Raw Component have
'   changed, when both of the above is true for a "Used Raw Component".
' ------------------------------------------------------------------------------
    Const PROC = "KindOfCodeChange_Get"
    
    On Error GoTo eh
    Dim fso                 As New FileSystemObject
    Dim sTempExport         As String
    Dim bCodeChanged        As Boolean
    Dim bRawChanged         As Boolean
    Dim bCloneChanged       As Boolean
    Dim sTempFolder         As String
    Dim arDiff              As Variant
    Dim flTempFormDesign    As FILE
    Dim flExpFormDesign     As FILE
    Dim flRawFormDesign     As FILE
    Dim flTempExport        As FILE
    
    If lKindOfCodeChange = enKindOfCodeChange.enUnknown Then
        
'        If cComp.CompName = "mFile" Then Stop
        '~~ Provide a temporary export for the current vbc which is used to determine a code change
        '~~ (a temporary folder is essential for the determination of a UserForm design change)
        With fso
            sTempFolder = sExpPath & "\Temp"
            If Not .FolderExists(sTempFolder) Then .CreateFolder sTempFolder
            sTempExport = sTempFolder & "\" & vbc.name & sExtension
            vbc.Export sTempExport
            Set flTempExport = .GetFile(sTempExport)
        End With
    
        Select Case Me.KindOfComp
            Case enRawClone
                '~~ The modified raw may have been re-imported already and thus only the export is pending
                If Not mFile.sDiffer(dif_file1:=flTempExport _
                                   , dif_file2:=cRaw.ExpFile _
                                    ) _
                And mFile.sDiffer(dif_file1:=flTempExport _
                                , dif_file2:=cComp.ExpFile _
                                  ) _
                Then
                    '~~ The code - yet not exported - is identical with the remote raw component's code (i.e. its Export File)
                    lKindOfCodeChange = enPendingExportOnly
                    cLog.Action = "Kind of CodeChange = " & KoCCStrng(lKindOfCodeChange)
                    GoTo xt
                End If
                '~~ 1. The (remote) raw has changed when the clone's Export File differs from the raw's Export File
                bRawChanged = _
                mFile.sDiffer(dif_file1:=Me.ExpFile _
                          , dif_file2:=cRaw.ExpFile _
                          , dif_ignore_empty_records:=ignore_empty_lines _
                          , dif_lines:=arDiff _
                           )
                 If Not bRawChanged And vbc.Type = vbext_ct_MSForm Then
                    '~~ If it is a UserForm the design may have changed only
                    Set flExpFormDesign = fso.GetFile(Replace(sExpFileFullName, ".frm", ".frx"))
                    Set flRawFormDesign = fso.GetFile(Replace(cRaw.ExpFileFullName, ".frm", ".frx"))
                    If UserFormSizeIsDifferent(sid_file_1:=flExpFormDesign _
                                             , sid_file_2:=flRawFormDesign _
                                              ) _
                    Then
                        bRawChanged = True
                    End If
                End If
                
                '~~ 2. The raw clone has changed when clone's temporary Export File differs from its last created Export File
                If Not fso.FileExists(sExpFileFullName) Then
                    bCloneChanged = True
                Else
                    bCloneChanged = _
                    mFile.sDiffer(dif_file1:=flTempExport _
                              , dif_file2:=Me.ExpFile _
                              , dif_ignore_empty_records:=ignore_empty_lines _
                              , dif_lines:=arDiff _
                               )
                End If
                    
                If Not bCloneChanged And vbc.Type = vbext_ct_MSForm Then
                    '~~ If it is a UserForm the design may have changed only
                    Set flExpFormDesign = fso.GetFile(Replace(sExpFileFullName, ".frm", ".frx"))
                    Set flTempFormDesign = fso.GetFile(Replace(sTempExport, ".frm", ".frx"))
                    If UserFormSizeIsDifferent(sid_file_1:=flTempFormDesign _
                                             , sid_file_2:=flExpFormDesign _
                                              ) _
                    Then
                        bCloneChanged = True
                    End If
                End If
                
            Case enHostedRaw, enInternal
                If Not fso.FileExists(sExpFileFullName) Then
                    bCodeChanged = True
                Else
                    bCodeChanged = mFile.sDiffer(dif_file1:=flTempExport _
                                             , dif_file2:=Me.ExpFile _
                                             , dif_ignore_empty_records:=False _
                                             , dif_lines:=arDiff _
                                              )
                End If
                
                If Not bCodeChanged And vbc.Type = vbext_ct_MSForm Then
                    '~~ If it is a UserForm the design may have changed only
                    Set flTempFormDesign = fso.GetFile(Replace(sTempExport, ".frm", ".frx"))
                    Set flExpFormDesign = fso.GetFile(Replace(sExpFileFullName, ".frm", ".frx"))
                    If UserFormSizeIsDifferent(sid_file_1:=flTempFormDesign _
                                             , sid_file_2:=flExpFormDesign _
                                              ) _
                    Then
                        bCodeChanged = True
                    End If
                End If
        End Select
            
        Select Case lKindOfComp
            Case enRawClone
                If bRawChanged And bCloneChanged Then
                    lKindOfCodeChange = enRawAndClone
                ElseIf bRawChanged And Not bCloneChanged Then
                    lKindOfCodeChange = enRawOnly
                ElseIf Not bRawChanged And bCloneChanged Then
                    lKindOfCodeChange = enCloneOnly
                ElseIf Not bRawChanged And Not bCloneChanged Then
                    lKindOfCodeChange = enNoCodeChange
                End If
            Case enHostedRaw, enInternal
                If bCodeChanged Then
                    lKindOfCodeChange = enInternalOnly
                Else
                    lKindOfCodeChange = enNoCodeChange
                End If
        End Select
        cLog.Action = "Kind of CodeChange = " & KoCCStrng(lKindOfCodeChange)
    End If
    
xt: KindOfCodeChange = lKindOfCodeChange
    If lKindOfCodeChange <> enNoCodeChange And lKindOfCodeChange <> enPendingExportOnly Then
        On Error Resume Next ' in case there is no line though a change had been detected (e.g. in the UserForm design only)
        cLog.Action = "1. Line changed: " & arDiff(0)
        On Error GoTo eh
    End If
    ' Do cleanup
    On Error Resume Next
    fso.DeleteFolder sTempFolder, True
    Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: Exit Property
    End Select
End Property

Private Function KoCCStrng(ByVal kocc As enKindOfCodeChange) As String
    Select Case kocc
        Case enKindOfCodeChange.enUnknown:              KoCCStrng = "Unknown (an error when logged!)"
        Case enKindOfCodeChange.enNoCodeChange:         KoCCStrng = "Unchanged"
        Case enKindOfCodeChange.enRawAndClone:          KoCCStrng = "Clone and Raw code changed"
        Case enKindOfCodeChange.enCloneOnly:            KoCCStrng = "Clone code changed"
        Case enKindOfCodeChange.enRawOnly:              KoCCStrng = "Raw code changed"
        Case enKindOfCodeChange.enInternalOnly:             KoCCStrng = "Code changed (is neither a raw nor a clone however!)"
        Case enKindOfCodeChange.enPendingExportOnly:    KoCCStrng = "Export pending only"
    End Select
End Function

Public Property Get KindOfComp() As enKindOfComp
    Const PROC = "KindOfComp"
    
    On Error GoTo eh
    
    mErH.BoP ErrSrc(PROC)
    If lKindOfComp = enKindOfComp.enUnknown Then
        If mRaw.Exists(sCompName) Then
            If mRaw.HostFullName(comp_name:=sCompName) = Me.Wrkbk.FullName _
            Then lKindOfComp = enHostedRaw _
            Else lKindOfComp = enRawClone
        ElseIf lKindOfComp = enKindOfComp.enUnknown Then
            lKindOfComp = enInternal
        End If
        cLog.Action = "Kind of Component = " & KoCStrng(lKindOfComp)
    End If
    KindOfComp = lKindOfComp
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property


Private Function KoCStrng(ByVal koc As enKindOfComp) As String
    Select Case koc
        Case enKindOfComp.enUnknown:        KoCStrng = "Unknown"
        Case enKindOfComp.enInternal:       KoCStrng = "Internal (neither raw nor clone)"
        Case enKindOfComp.enRawClone:       KoCStrng = "Raw Clone"
        Case enKindOfComp.enHostedRaw:      KoCStrng = "Hosted Raw"
    End Select
End Function

Private Property Get ReplaceRemoteWithClonedRawConfirmed() As Boolean
' -------------------------------------------------------------
' Returns TRUE when the "reverse" code update is confirmed.
' --------------------------------------------------------------
    Const PROC = "ReplaceRemoteWithClonedRawConfirmed"
    
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
                         " with this component's Export File. I.e. the changes" & vbLf & _
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
        .section(1).sLabel = "About this incident:"
        .section(1).sText = "The module '" & cComp.CompName & "' is a clone of the raw component hosted in Workbook '" & cRaw.WrkbkBaseName & "' " & _
                            "This clone component's code has usually has been synched with the remote raw componet's code. " & _
                            "The remote raw hasn't been changed but the cloned component's code has."
        .section(2).sLabel = "About the choices:"
        .section(2).sText = "When the code modification is ignored it will be regarded temporarily only and overwrtitten with the next open of this Workbook. " & _
                            "When the remote The module, updated when this Workbook has opened, is no longer identical with its origin. " & _
                            "I.e. the component's code has been changed in this Workbook using it instead of in the " & _
                            "Workbook the origin is ""hosted""."
        .section(3).sText = "Take your decission or postpone it after having displayed the difference."
    End With
        
    sReply = sReplyDsplyDiff
    Do
        sReply = mMsg.Dsply(msg_title:=sTitle, _
                            msg:=sMsg, _
                            msg_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sReplyUpdateOrigin
                ReplaceRemoteWithClonedRawConfirmed = True
                Exit Do
            Case sReplyDsplyDiff
                mFile.Compare file_left_full_name:=cComp.ExpFileFullName _
                          , file_right_full_name:=cRaw.ExpFileFullName _
                          , file_left_title:="The cloned raw's current code in Workbook/VBProject " & cComp.WrkbkBaseName & " (" & cComp.ExpFileFullName & ")" _
                          , file_right_title:="The remote raw's current code in Workbook/VBProject " & cRaw.WrkbkBaseName & " (" & cRaw.ExpFileFullName & ")"
                ' The question is re-displayed for a decison
            Case sReplyIgnore
                ReplaceRemoteWithClonedRawConfirmed = False
                Exit Do
        End Select
    Loop
                                            
xt: Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Public Property Get UpdatedAsOf() As Date:                      UpdatedAsOf = dtUpdatedAsOf:            End Property

Public Property Let UpdatedAsOf(ByVal dt As Date):              dtUpdatedAsOf = dt:                     End Property

Public Property Get VBComp() As VBComponent
    Const PROC = "VBComp_Get"
    
    On Error GoTo eh
    
    If vbc Is Nothing Then
        If Not mCompMan.CompExists(ce_wb:=wb, ce_comp_name:=sCompName) _
        Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "A Component named '" & sCompName & "' does not exist in Workbook '" & wb.name & "'!"
        Set vbc = wb.VBProject.VBComponents(sCompName)
    End If
    Set VBComp = vbc

xt: Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Public Property Let VBComp(ByVal vb_comp As VBComponent):       Set vbc = vb_comp:                      End Property

Public Property Get Wrkbk() As Workbook
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
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Public Property Let Wrkbk(ByVal wbk As Workbook)
' ----------------------------------------------
' Provides the properties:
' - Wrkbk (wb),
' - WrkbFullName (sWrkbkFullName)
' - WrkbkBaseName (sWrkbkBaseName)
' ---------------------------------------------
    
    Dim fso As New FileSystemObject
    
    Set wb = wbk
    sWrkbkFullName = wb.FullName
    sWrkbkBaseName = fso.GetBaseName(sWrkbkFullName)
    Set foExport = fso.GetFolder(Me.ExpPath)
    For Each vbc In wb.VBProject.VBComponents
        lCompMaxLen = mBasic.Max(lCompMaxLen, Len(vbc.name))
    Next vbc
    Set fso = Nothing
    
End Property

Public Property Get WrkbkBaseName() As String:                 WrkbkBaseName = BaseName(sWrkbkFullName):        End Property

Public Property Get WrkbkFullName() As String:                 WrkbkFullName = sWrkbkFullName:                  End Property

Public Property Let WrkbkFullName(ByVal s As String):          sWrkbkFullName = s:                              End Property

Public Sub BackUpCode()
' -------------------------------------------------------------------------
' Exports the VBComponent (vbc) provided the CodeModule is not empty.
' Note: The component is exported even when the code had not changed.
' -------------------------------------------------------------------------
    Const PROC = "BackUpCode"
    
    On Error GoTo eh
    
    If Not vbc.CodeModule.CountOfLines = 0 Then
        vbc.Export sExpFileFullName
        Set flExpFile = GetFile(sExpFileFullName)
        cLog.Action = "Exported to '" & sExpFileFullName & "'"
    End If

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Function BaseName(ByVal s As String) As String
    Dim fso As New FileSystemObject
    With fso:  BaseName = .GetBaseName(s): End With
    Set fso = Nothing
End Function

Public Function CodeModuleIsEmpty() As Boolean
' --------------------------------------------
' Returns TRUE when the CodeModule (v) has 0
' only 1 line with a lenght < 2.
' --------------------------------------------
    
    With vbc.CodeModule
        If .CountOfLines = 0 Then
            CodeModuleIsEmpty = True
        ElseIf .CountOfLines = 1 And Len(.Lines(1, 1)) < 2 Then
            CodeModuleIsEmpty = True
        End If
    End With
    
End Function

Public Function CodeToArray() As String()
' ---------------------------------------
' Returns the Component's CodeModule as
' as Array. When the Code Module is empty
' The CodeArray will not be assigned !!!!
' ---------------------------------------
    Const PROC = "CodeToArray"

    On Error GoTo eh
    Dim va      As Variant
    Dim sLine   As String
    Dim sSplit  As String
        
    With vbc.CodeModule
        If .CountOfLines = 0 Then GoTo xt
        sLine = .Lines(1, .CountOfLines)
        If InStr(sLine, vbCr) <> 0 Then sSplit = vbCr
        If InStr(sLine, vbLf) <> 0 Then sSplit = sSplit & vbLf
        va = Split(sLine, sSplit)
    End With
            
    mBasic.ArrayTrimm va
    
xt: If Not TypeName(va) = "Empty" _
    Then CodeToArray = va
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Public Function CodeToDict() As Dictionary
' ----------------------------------------
' Returns the Component's CodeModule as
' Dictionary.
' ---------------------------------------
    Const PROC = "CodeToDict"

    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim va      As Variant
    Dim sLine   As String
    Dim sSplit  As String
    Dim i       As Long
    Dim v       As Variant
        
    With vbc.CodeModule
        If .CountOfLines = 0 Then GoTo xt
        sLine = .Lines(1, .CountOfLines)
        If InStr(sLine, vbCr) <> 0 Then sSplit = vbCr
        If InStr(sLine, vbLf) <> 0 Then sSplit = sSplit & vbLf
        va = Split(sLine, sSplit)
    End With
    mBasic.ArrayTrimm va
    For Each v In va
        i = i + 1
        dct.Add i, v
    Next v
    
xt: Set CodeToDict = dct
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function DebugComp() As String
' ------------------------------------
'
' ------------------------------------
        
    If lMaxCompLength > 0 _
    Then DebugComp = sCompName & " " & String(lMaxCompLength - Len(sCompName), ".") & ": " _
    Else DebugComp = sCompName & " : "

End Function

Public Sub DisplayDiff(ByVal flLeft As FILE, _
                       ByVal flRght As FILE, _
              Optional ByVal sTitleLeft As String, _
              Optional ByVal sTitleRght As String)
' ------------------------------------------------
' Displays the difference between file 1 (fl1) and
' file 2 (fl2).
' ------------------------------------------------
    Const PROC = "DisplayDiff"
    
    Dim waitOnReturn    As Boolean: waitOnReturn = True
    Dim windowStyle     As Integer: windowStyle = 1
    Dim sCommand        As String

    If Not mBasic.AppIsInstalled("WinMerge") _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "Display of the difference between two files relies on WinMerge installed, which is not the case!"
    
    sCommand = "WinMergeU /e" & _
               " /dl " & DQUOTE & sTitleLeft & DQUOTE & _
               " /dr " & DQUOTE & sTitleRght & DQUOTE & " " & _
               """" & flLeft.PATH & """" & " " & _
               """" & flRght.PATH & """"
    
    With New wshShell
        .Run sCommand, windowStyle, waitOnReturn
    End With
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComp" & "." & sProc
End Function

Public Function Exists(ByVal sComp As String) As Boolean
    Dim vbc As VBComponent

    On Error Resume Next
    Set vbc = Me.Wrkbk.VBProject.VBComponents(sComp)
    Exists = Err.Number = 0
    Set vbc = Nothing

End Function

Public Property Get ExpPath() As String
' -----------------------------------------
' Returns the path for the code backup
' export file  of the Component which
' defaults to the Workbook's path with a
' Workbook's base name folder.
' Will be created if not existent.
' --------------------------------------
    Const PROC = "ExpPath_Get"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim sPath   As String
    
    If sExpPath = vbNullString Then
        With fso
            If Right(.GetParentFolderName(wb.FullName), Len(sWrkbkBaseName)) <> sWrkbkBaseName Then
               sPath = .GetParentFolderName(wb.FullName) & "\" & sWrkbkBaseName
                '~~ Make sure the returned export path specifies an existing folder
                If Not .FolderExists(sPath) Then
                    .CreateFolder (sPath)
                End If
            Else
                sPath = .GetParentFolderName(wb.FullName)
            End If
            sExpPath = sPath
        End With
    End If
    
xt: ExpPath = sExpPath
    Set fso = Nothing
    Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Public Function ExpFileToArray(ByVal fl As FILE) As String()
' ---------------------------------------------------------------------
' Returns the content of the file (vFile) which may be provided as file
' object or full file name string as array, considering any kind of
' line break characters.
' ---------------------------------------------------------------------
    Const PROC = "ExpFileToArray"
    Const ATTRIB_VB = "Attribute VB_"

    On Error GoTo eh
    Dim a           As Variant
    Dim i           As Long
    
    a = mFile.ToArray(fl)
    '~~ Remove the extra module information lines
    For i = UBound(a) To LBound(a) Step -1
        '~~ Find last Exported module info
        If left(a(i), Len(ATTRIB_VB)) = ATTRIB_VB Then
            Exit For
        End If
    Next i
    If i >= LBound(a) Then mBasic.ArrayRemoveItems a, Element:=1, NoOfElements:=i + 1
    
    '~~ Remove any leading or trailing empty items
    If mBasic.ArrayIsAllocated(a) Then
        mBasic.ArrayTrimm a
        If a(UBound(a)) = vbNullString Then Stop
    End If
    
     ' remove leading and trailing empty items
    mBasic.ArrayTrimm a
    ExpFileToArray = a

xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Public Function ExpFileToDict(ByVal fl As FILE) As Dictionary
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
    Set dct = mFile.ToDict(fl)
    
    '~~ Remove the extra module information lines
    For Each v In dct
        If left(sLine, Len(ATTRIB_LINE)) = ATTRIB_LINE _
        And left(dct(v), Len(ATTRIB_LINE)) <> ATTRIB_LINE _
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
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function GetFile(ByVal sFile As String) As FILE
    Dim fso As New FileSystemObject
    With fso:  Set GetFile = .GetFile(sFile):  End With
    Set fso = Nothing
End Function

Public Sub RenewByImport(ByVal rn_wb As Workbook, _
                 ByVal rn_comp_name As String, _
                 ByVal rn_exp_file_full_name As String, _
        Optional ByRef rn_log As clsLog)
' --------------------------------------------------------
' Renews the component (rn_comp_name) in Workbook (rn_wb)
' by re-importing the Export File (rn_exp_file_full_name).
' - When the Workbook (wb) is the ActiveWorkbook the
'   service is aborted without notice. I.e. de-activating
'   the active Workbook is with the caller
' - When the base name of the provided Export File is not
'   identical with the current component's name
'   (rn_comp_name) the service is aborted without notice.
' - When the component exists but is a Data Module type
'   the service is terminated without notic
' - Logging of the process is recommended in order to see
'   at which state it may have failed.
'
' W. Rauschenberger Berlin, Jan 2021
' --------------------------------------------------------
    Const PROC = "RenewByImport"
    
    On Error GoTo eh
    Dim sTempName   As String
    Dim i           As Long
    Dim j           As Long
    Dim fso         As New FileSystemObject
    Dim vbc         As VBComponent
    
    '~~ Preconditions
    If rn_wb Is ActiveWorkbook Then GoTo xt
    If fso.GetBaseName(rn_exp_file_full_name) <> rn_comp_name Then GoTo xt
    If Not mCompMan.CompExists(ce_wb:=rn_wb, ce_comp_name:=rn_comp_name) Then GoTo xt
    If rn_wb.VBProject.VBComponents(rn_comp_name).Type = vbext_ct_Document Then GoTo xt
    
    sTempName = AvailableCompTempName(ac_vbp:=rn_wb.VBProject, ac_comp_name:=rn_comp_name)
    With rn_wb.VBProject
        ' --------------------------------------------------------------------
        '~~ 1. Step: Rename and remove the component (though the component may
        '~~          not be removed before this this service has ended it will
        '~~          be out of the way for the re-import of the Export File.
        '~~ ------------------------------------------------------------------
        On Error Resume Next
        .VBComponents(rn_comp_name).name = sTempName
        If Err.Number <> 0 Then
            '~~ When the component cannot be renamed it is removed and the import
            '~~ is postponed to the mPending.Resolve service
            If Not rn_log Is Nothing Then rn_log.Action = "Renaming component to '" & sTempName & "' failed with '" & Err.Description & "' (err-no = " & Err.Number
            mPending.Serviced = rn_wb
            mPending.Import(im_wb:=rn_wb _
                          , im_exp_file_full_name:=rn_exp_file_full_name _
                          , im_comp_type:=.VBComponents(rn_comp_name).Type _
                           ) = rn_comp_name
            If Not rn_log Is Nothing Then rn_log.Action = "Renew postponed, i.e. import of '" & rn_exp_file_full_name & "' is pending"
            
            On Error Resume Next
            .VBComponents.Remove .VBComponents(rn_comp_name)
            If Err.Number <> 0 Then
                If Not rn_log Is Nothing Then rn_log.Action = "Removal failed with  '" & Err.Description & "' (err-no = " & Err.Number
            Else
                If Not rn_log Is Nothing Then rn_log.Action = "Component '" & rn_comp_name & "' removed"
            End If
            
            If mCompMan.CompExists(ce_wb:=rn_wb, ce_comp_name:=sTempName) Then
                On Error Resume Next
                .VBComponents.Remove .VBComponents(sTempName)
                If Err.Number <> 0 Then
                    If Not rn_log Is Nothing Then rn_log.Action = "Removal of '" & sTempName & "' failed with  '" & Err.Description & "' (err-no = " & Err.Number
                Else
                    If Not rn_log Is Nothing Then rn_log.Action = "Component '" & sTempName & "' removed"
                End If
            End If
            GoTo xt
        Else
            If Not rn_log Is Nothing Then rn_log.Action = "Component renamed to '" & sTempName & "'"
            On Error Resume Next
            .VBComponents.Remove .VBComponents(sTempName) ' will take place when this service has ended
            If Err.Number <> 0 Then
                If Not rn_log Is Nothing Then rn_log.Action = "Removal of '" & sTempName & "' failed with '" & Err.Description & "' (err-no = " & Err.Number
            Else
                If Not rn_log Is Nothing Then rn_log.Action = "Component '" & sTempName & "' removed"
            End If
        End If
        
        '~~ ----------------------------------
        '~~ 2. Step: Re-Import the Export File
        '~~ ----------------------------------
        RenewWaitAndSaveWbk rs_wb:=rn_wb, rs_log:=rn_log, rs_wait_msec:=5
        On Error Resume Next
        .VBComponents.Import rn_exp_file_full_name
        If Err.Number <> 0 Then
            If Not rn_log Is Nothing Then rn_log.Action = "Renew by import of Export File '" & rn_exp_file_full_name & "' failed with '" & Err.Description & "' (err-no = " & Err.Number
        Else
            If Not rn_log Is Nothing Then rn_log.Action = "Component renewed by re-import of Export File '" & rn_exp_file_full_name & "'"
        End If
        RenewWaitAndSaveWbk rs_wb:=rn_wb, rs_log:=rn_log, rs_wait_msec:=5
        
        '~~ ------------------------------------------------------------------
        '~~ 3. Step: Re-export the component in case the provided Export File
        '~~          was not identical with the component's origin Export File
        '~~ ------------------------------------------------------------------
        If rn_exp_file_full_name <> sExpFileFullName Then
            .VBComponents(rn_comp_name).Export sExpFileFullName
        End If
        
    End With

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub RenewWaitAndSaveWbk( _
                          ByRef rs_wb As Workbook, _
                 Optional ByVal rs_wait_msec As Long = 0, _
                 Optional ByRef rs_log As clsLog)
' ---------------------------------------------------------
'
' ---------------------------------------------------------
    DoEvents
    Application.Wait Now() + (rs_wait_msec * 0.00000001)
    Application.EnableEvents = False
    On Error Resume Next
    rs_wb.Save
    If Err.Number <> 0 Then
        If Not rs_log Is Nothing Then rs_log.Action = "Workbook save failed with '" & Err.Description & "' (err-no = " & Err.Number
    Else
        If Not rs_log Is Nothing Then rs_log.Action = "Waited for " & rs_wait_msec & " milliseconds and Workbook saved (with EnableEvents = False)"
    End If
    Application.EnableEvents = True
    
End Sub

Private Function AvailableCompTempName( _
                 ByVal ac_vbp As VBProject, _
                 ByVal ac_comp_name As String) As String
' -------------------------------------------------
    Const PROC = ""
    
    On Error GoTo eh
    Dim sTempName   As String
    Dim i           As Long
    Dim j           As Long
    
    sTempName = ac_comp_name & "_Temp"
    With ac_vbp
        Do
            '~~ When the primarily set temp name for the to-be-replace component is already existing
            '~~ a yet non-existing temp name has to be found
            For i = 1 To .VBComponents.Count
                If .VBComponents(i).name = sTempName Then
                    .VBComponents.Remove .VBComponents(sTempName) ' left over from any previous aborted service
                    j = j + 1: sTempName = sTempName & j
                    GoTo lp
                End If
            Next i
lp:         If i > .VBComponents.Count Then Exit Do
        Loop
    End With
    
xt: AvailableCompTempName = sTempName
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Public Sub ReplaceRemoteWithClonedRaw( _
           ByVal raw_remote_comp_name As String, _
           ByRef rr_log As clsLog)
' -------------------------------------------------
' Replace the raw component's code with the Export
' File of the clone component.
' -------------------------------------------------
    Const PROC = "ReplaceRemoteWithClonedRaw"

    On Error GoTo eh
    Dim wbRemoteRaw As Workbook
    Dim vbcRaw      As VBComponent
    Dim sTempName   As String: sTempName = raw_remote_comp_name & "_Temp"
    Dim i           As Long
    
    '~~ When the primary set temp name for the to-be-replace vbc is already existing
    '~~ a yet non-existing temp name has to be found because an existing vbc cannot
    '~~ be removed beforehand but only afterwards
    While mCompMan.CompExists(ce_wb:=wb, ce_comp_name:=sTempName)
        i = i + 1: sTempName = sTempName & i
    Wend
    
    Set wbRemoteRaw = mCompMan.WbkGetOpen(cRaw.HostFullName) ' Wrkbk is the Workbook hosting the raw Component
    Set vbcRaw = wbRemoteRaw.VBProject.VBComponents(sCompName)
    
    Me.RenewByImport rn_wb:=wbRemoteRaw _
                      , rn_comp_name:=sCompName _
                      , rn_exp_file_full_name:=sExpFileFullName
    cLog.Action = "Remote raw replaced by re-import of the clone's export file '" & sExpFileFullName & "'"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub ReplaceRemoteWithClonedRawWhenConfirmed( _
           ByRef rwu_updated As Boolean, _
           ByRef rwu_log As clsLog)
' ------------------------------------------------------------------------
' Provided confirmed by the user, replaces the raw code in the hosting
' Workbook with the cloned raw component's code in its Export File.
' Code replacement for class modules, satandard modules and UserForms is
' done via a rename, re-import, remove renamed procedure, for a data
' module the code is replaced line by line because a (re-)import is
' impossible for such a module.
' ------------------------------------------------------------------------
    Const PROC = "ReplaceRemoteWithClonedRawWhenConfirmed"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If ReplaceRemoteWithClonedRawConfirmed Then
            '~~ Get the origin host Workbook open
        With cRaw
            .HostFullName = mRaw.HostFullName(comp_name:=cComp.CompName)
            .ExpFile = fso.GetFile(mRaw.ExpFileFullName(comp_name:=cComp.CompName))
            .ExpFileFullName = .ExpFile.PATH
            .RawHost = mCompMan.WbkGetOpen(.HostFullName)
            .RawVbc = .RawHost.VBProject.VBComponents(Index:=cComp.CompName)
            Select Case .RawType
                Case vbext_ct_Document
                    '~~ A data module cannot be replaced like the other ones. Thus only its code
                    '~~ is synchronized - disregarding any consequences this may have when the
                    '~~ structure and design of the underlying Workbook or Worksheet has changed.
                    cComp.SynchRemoteWithClonedRaw
                Case vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_MSForm
                    Me.ReplaceRemoteWithClonedRaw raw_remote_comp_name:=.CompName, rr_log:=rwu_log
            End Select
            rwu_updated = True
        End With ' cRaw
    End If

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SynchClonedWithRemoteRaw()
' ----------------------------------------------------
' Synchronizes cloned raw component's code with
' the remote raws component's code in its Export File.
' ----------------------------------------------------
    Const PROC = "SynchClonedWithRemoteRaw"

    On Error GoTo eh
    Dim iLine   As Long
    Dim v       As Variant
    Dim vbcm    As CodeModule
    Dim fso     As New FileSystemObject
    
    '~~ Remove all lines from the cloned raw component
    Set vbcm = vbc.CodeModule
    With vbcm
        .DeleteLines 1, .CountOfLines
    
        '~~ Insert/copy all code lines from the source to the target component
        iLine = 0
        For Each v In ExpFileToArray(cRaw.ExpFile)
            iLine = iLine + 1
            .InsertLines iLine, v
        Next v
    End With
    
    Debug.Print DebugComp & Now() & " Code of the cloned raw component's CodeModule synchronized with code the remote raw component's Export File '" & cRaw.ExpFileFullName & "'."
                
xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SynchRemoteWithClonedRaw()
' ------------------------------
' Synchronizes the Common
' Component's origin code with
' the Export File of the used
' Common Component.
' ------------------------------
    Const PROC = "SynchRemoteWithClonedRaw"

    On Error GoTo eh
    Dim iLine   As Long
    Dim v       As Variant
    Dim vbcm    As CodeModule
    Dim fso     As New FileSystemObject
                
    Debug.Print DebugComp & Now() & " About to synch the remote raw component's code with the code in the cloned raw's code in the Export File '" & cComp.ExpFileFullName & "'."
    '~~ Remove all lines from the target component
    Set vbcm = cRaw.RawVbc.CodeModule
    With vbcm
        .DeleteLines 1, .CountOfLines
        '~~ Insert all code lines from the cloned raw component's Export File into the remote raws CodeModule
        iLine = 0
        For Each v In ExpFileToArray(cComp.ExpFile)
            iLine = iLine + 1
            vbcm.InsertLines iLine, v
        Next v
    End With
    
    Debug.Print DebugComp & Now() & " Code of remote raw component synchronized with the code in the cloned raw's Export File '" & cComp.ExpFileFullName & "'."
                
xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Function TypeString() As String
' --------------------------------------------------------------------------
' Returns the component type derived from sComp.
' --------------------------------------------------------------------------
    Const PROC = "TypeString"
    
    On Error GoTo eh
        
    Select Case vbc.Type
        Case vbext_ct_ActiveXDesigner:  TypeString = "ActiveX Designer"
        Case vbext_ct_ClassModule:      TypeString = "Class Module"
        Case vbext_ct_Document:         TypeString = "Document Module"
        Case vbext_ct_MSForm:           TypeString = "UserForm"
        Case vbext_ct_StdModule:        TypeString = "Code Module"
        Case Else:                      TypeString = "Type '" & CStr(vbc.Type) & "' unknown"
    End Select
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function UserFormSizeIsDifferent( _
                                   ByVal sid_file_1 As FILE, _
                                   ByVal sid_file_2 As FILE) As Boolean
    Dim sFrxFile1Path  As String
    Dim sFrxFile2Path  As String
    Dim fso             As New FileSystemObject
    
    sFrxFile1Path = Replace$(sid_file_1.PATH, ".frm", ".frx")
    sFrxFile2Path = Replace$(sid_file_2.PATH, ".frm", ".frx")
    
    With fso
        If sid_file_1.Size <> sid_file_2.Size Then
            UserFormSizeIsDifferent = True
        ElseIf .GetFile(sFrxFile1Path).Size <> .GetFile(sFrxFile2Path).Size Then
            UserFormSizeIsDifferent = True
        End If
    End With
    Set fso = Nothing
    
End Function

