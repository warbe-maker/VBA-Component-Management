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
' - ExportFile              R/W File object created by menas of the BackUpCode method
' - HostedRawComps  W   String with named Components hosted in the Workbook
' - IsRaw                R   TRUE when a Component is regarded a Common one
' - CodeChanged             R   TRUE when the code of the CodeModule differs from the code
'                               backed up in the corresponding ExportFile
' - KindOfChange            R
' - KindOfComp              R
'
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
Private flExport            As FILE
Private foExport            As Folder
Private sPathCommon         As String
Private sExportFullName     As String       ' Full filespec for the Export File created by the method BackUpCode
Private lCompMaxLen         As Long
Private lKindOfComp         As enKindOfComp
Private lKindOfCodeChange   As enKindOfCodeChange
'Private sRawExpFileName     As String
'Private sRawHostFullName    As String
Private dctHostedRaws       As New Dictionary

Private Sub Class_Initialize()
    sPathCommon = mCfg.CommonComponentsBasePath
    lCompMaxLen = mDat.CommCompsMaxLenght()
    Set dctRaws = mDat.RawComps
    Set dctRawHosts = mDat.HostWorkbooks
End Sub

'Public Property Let RawExpFileName(ByVal fl_name As String):                sRawExpFileName = fl_name:              End Property
'Public Property Get RawExpFileName() As String:                            RawExpFileName = sRawExpFileName:       End Property
'Public Property Let RawHostFullName(ByVal fl_name As String):               sRawHostFullName = fl_name:             End Property
'Public Property Get RawHostFullName() As String:                           RawHostFullName = sRawHostFullName:     End Property

Public Property Get CodeVersionAsOfDate() As Date
    CodeVersionAsOfDate = mDat.CodeVersionAsOfDate(sWrkbkBaseName, sCompName)
End Property

Public Property Let CodeVersionAsOfDate(ByVal dtAsOf As Date):
    mDat.CodeVersionAsOfDate(sWrkbkBaseName, sCompName) = dtAsOf
End Property

Public Property Get CompName() As String:                       CompName = sCompName:                   End Property

Public Property Let CompName(ByVal s As String):                sCompName = s:                          End Property

Public Property Get ComponentName() As String:                  ComponentName = sCompName:              End Property

Public Property Let ComponentName(ByVal s As String):           sCompName = s:                          End Property

Public Property Get ExportFile() As FILE
    
    Dim sFile As String

    If flExport Is Nothing Then
        sFile = ExportFileFullName
        If FileExists(sFile) Then
            Set flExport = GetFile(sFile)
        Else
            '~~ For the very first check of a code change there code will not have been backe up
            '~~ Backing up the code will result in "code not changed" which is ok since the code
            '~~ had already been backed up.
            Me.BackUpCode
        End If
    End If
    
    Set ExportFile = flExport

End Property

Public Property Let ExportFile(ByVal fl As FILE):               Set flExport = fl:                      End Property

Private Property Get ExportFileExtention() As String
    Select Case vbc.Type
        Case vbext_ct_StdModule:    ExportFileExtention = ".bas"
        Case vbext_ct_ClassModule:  ExportFileExtention = ".cls"
        Case vbext_ct_MSForm:       ExportFileExtention = ".frm"
        Case vbext_ct_Document:     ExportFileExtention = ".cls"
    End Select
End Property

Public Property Get ExportFileFullName() As String
    ExportFileFullName = sExportFullName
End Property

Public Property Get ExportFolder() As Folder:               Set ExportFolder = foExport:                End Property

Public Property Get Extension() As String:                  Extension = TypeExtention(vbc):             End Property

Private Property Get IsHostedRaw(Optional ByVal comp_name As String) As Boolean
    IsHostedRaw = HostedRawComps.Exists(comp_name)
End Property

Private Property Get HostedRawComps(Optional ByVal hosted_raws As String) As Variant
    Set HostedRawComps = dctHostedRaws
End Property

Public Property Let HostedRawComps( _
           Optional ByVal hosted_raws As String, _
                    ByVal vhosted As Variant)
' -------------------------------------------
' Saves the name of raw components
' (hosted_raws) to the Dictionary (dctRaws).
' -------------------------------------------
    Dim v               As Variant
    Dim sComp           As String
    
    For Each v In Split(hosted_raws, ",")
        sComp = Trim$(v)
        If Not dctHostedRaws.Exists(sComp) Then dctHostedRaws.Add sComp, sComp
    Next v
    If dctHostedRaws.Count <> 0 Then
        mDat.CommCompsHostWorkbookFullName(BaseName(wb.name)) = wb.FullName
    End If
    
End Property

Public Property Get IsRaw() As Boolean
' ---------------------------------------------------------------
' IsRaw is the oposite of IsHostedRaw. When the Component is a
' Common used one the Export File is returned as object.
' ---------------------------------------------------------------
    IsRaw = mDat.IsRaw(raw_comp_name:=sCompName, raw_exp_file:=Me.ExportFile, raw_host_full_name:=Me.WrkbkFullName)

End Property

Public Property Get KindOfCodeChange() As enKindOfCodeChange:               KindOfCodeChange = lKindOfCodeChange:   End Property
Public Property Let KindOfCodeChange(ByVal kocc As enKindOfCodeChange):     lKindOfCodeChange = kocc:               End Property

Public Property Get KindOfComp() As enKindOfComp
    Const PROC = "KindOfComp"
    
    On Error GoTo eh
    
    If lKindOfComp = enUnknown Then
        If IsHostedRaw(Me.CompName) Then
            lKindOfComp = enRemoteRaw
            GoTo xt
        End If
        
        If dctRaws.Exists(Me.CompName) And mDat.CommCompHostWorkbookBaseName(Me.CompName) = Me.WrkbkBaseName Then
            lKindOfComp = enRemoteRaw
            GoTo xt
        End If
        
        If dctRaws.Exists(Me.CompName) And mDat.CommCompHostWorkbookBaseName(Me.CompName) <> Me.WrkbkBaseName Then
            lKindOfComp = enClonedRaw
            GoTo xt
        End If
        
        If lKindOfComp <> enRemoteRaw And lKindOfComp <> enClonedRaw Then
            lKindOfComp = enNoRaw
            GoTo xt
        End If
    Else
        KindOfComp = lKindOfComp
    End If
    
    Select Case lKindOfComp
        Case enRemoteRaw
            Me.RegisterAsHostWorkbook
            mDat.CommCompHostWorkbookBaseName(sComp:=Me.CompName) = Me.WrkbkBaseName
            mDat.CommCompExpFileFullName(sComp:=Me.CompName) = Me.ExportFileFullName
        Case enClonedRaw
'            Me.RegisterAsCloneWorkbook
    End Select
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Public Property Get UpdatedAsOf() As Date:                      UpdatedAsOf = dtUpdatedAsOf:            End Property

Public Property Let UpdatedAsOf(ByVal dt As Date):              dtUpdatedAsOf = dt:                     End Property

Private Property Get ReplaceRemoteWithClonedRawConfirmed() As Boolean
' -------------------------------------------------------------
' Returns TRUE when the "reverse" code update is confirmed.
' --------------------------------------------------------------
    Const PROC = "ReplaceRemoteWithClonedRawConfirmed"
    
    On Error GoTo eh
    Dim sReplyUpdateOrigin      As String
    Dim sReplyDsplyDiff         As String
    Dim sReplyIgnore            As String
    Dim sBaseMsg                As String
    Dim sMsg                    As tMsg
    Dim sDsplyDiffMsg           As String
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
        .Section(1).sLabel = "About this incident:"
        .Section(1).sText = "The module '" & cComp.CompName & "' is a clone of the raw component hosted in Workbook '" & cRemoteRaw.WrkbkBaseName & "' " & _
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
        sReply = mMsg.Dsply(dsply_title:=sTitle, _
                            dsply_msg:=sMsg, _
                            dsply_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sReplyUpdateOrigin
                ReplaceRemoteWithClonedRawConfirmed = True
                Exit Do
            Case sReplyDsplyDiff
                mFl.Compare file_left_full_name:=cComp.ExportFileFullName _
                          , file_right_full_name:=cRemoteRaw.ExportFileFullName _
                          , file_left_title:="The cloned raw's current code in Workbook/VBProject " & cComp.WrkbkBaseName & " (" & cComp.ExportFileFullName & ")" _
                          , file_right_title:="The remote raw's current code in Workbook/VBProject " & cRemoteRaw.WrkbkBaseName & " (" & cRemoteRaw.ExportFileFullName & ")"
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
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Public Property Get VBComp() As VBComponent
' ------------------------------------------------------------------------------
' ------------------------------------------------------------------------------
    If vbc Is Nothing Then
        If wb Is Nothing Then
            '~~ When there is no VBComponent object this indicates one is required for
            '~~ a remote raw component in order to update it with the cloned raw's
            '~~ Export File. From the remote raw component only its name and the
            '~~ FullName of the Workbook it is hosted is known.
            Set wb = mWrkbk.GetOpen(cRemoteRaw.WrkbkFullName)
        End If
        Set vbc = wb.VBProject.VBComponents(cComp.CompName)
    End If
    Set VBComp = vbc
End Property

Public Property Let VBComp(ByVal vb_comp As VBComponent)
    Const PROC = "VBComp_Let"

    On Error GoTo eh
    If Me.Wrkbk Is Nothing _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "There's yet no Wrkbk Workbook assigned!"
    
    Set vbc = vb_comp
    Me.CompName = vbc.name
    sExportFullName = ExportFilePath & "\" & Me.CompName & ExportFileExtention
    If Not FileExists(sExportFullName) Then
        Me.BackUpCode
    End If
    If FileExists(sExportFullName) _
    Then Me.ExportFile = GetFile(sExportFullName)
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Public Property Get Wrkbk() As Workbook
    If wb Is Nothing _
    Then Set wb = mWrkbk.GetOpen(sWrkbkFullName) _
    Else Set Wrkbk = wb
End Property

Public Property Let Wrkbk(ByVal wbk As Workbook)
    
    Dim fso As New FileSystemObject
    
    Set wb = wbk
    sWrkbkFullName = wb.FullName
    sWrkbkBaseName = fso.GetBaseName(sWrkbkFullName)
    Set foExport = fso.GetFolder(ExportFilePath)
    For Each vbc In wb.VBProject.VBComponents
        lCompMaxLen = mBasic.Max(lCompMaxLen, Len(vbc.name))
    Next vbc

End Property

Public Property Get WrkbkBaseName() As String:                 WrkbkBaseName = BaseName(sWrkbkFullName):        End Property

Public Property Get WrkbkFullName() As String:                 WrkbkFullName = sWrkbkFullName:                  End Property

Public Property Let WrkbkFullName(ByVal s As String):          sWrkbkFullName = s:                              End Property

Public Sub BackUpCode()
' -------------------------------------------------------------------------
' Exports the VBComponent (Me.VBComp) provided the CodeModule is not empty.
' Note: The component is exported even when the code had not changed.
' -------------------------------------------------------------------------
    Const PROC = "BackUpCode"
    On Error GoTo eh
    
    With Me
        If Not .VBComp.CodeModule.CountOfLines = 0 Then
            .VBComp.Export .ExportFileFullName
            mCompManLog.LogAction log_wb:=.WrkbkBaseName _
                                , log_action:=DebugComp & "Code change backed-up (exported)"
            .ExportFile = GetFile(sExportFullName)
        End If
    End With

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function BaseName(ByVal s As String) As String
    With New FileSystemObject:  BaseName = .GetBaseName(s): End With
End Function

Public Function CodeChanged( _
             Optional ByVal raw_remote As clsComp, _
             Optional ByVal raw_clone As clsComp, _
             Optional ByVal ignore_empty_lines As Boolean = False) As Boolean
' -------------------------------------------------------------------------
' Returns TRUE the Component's code either has never been exported of
' when it differs from its current Export File. A temporary Export File
' is used for the comparison of the current Component's code with its
' Export File in order to have two files to compare.
' -------------------------------------------------------------------------
    Const PROC  As String = "CodeChanged"
       
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim sTempExp    As String
    
    Select Case lKindOfComp
        Case enClonedRaw, enNoRaw
            sTempExp = fso.GetParentFolderName(Me.ExportFileFullName) & "\Temp_" & fso.GetFileName(Me.ExportFileFullName)
            Me.VBComp.Export sTempExp
            Select Case Me.VBComp.Type
                Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document
                    CodeChanged = mFl.sDiffer(dif_file1:=fso.GetFile(sTempExp) _
                                            , dif_file2:=Me.ExportFile _
                                            , dif_ignore_empty_records:=ignore_empty_lines _
                                             )
                Case vbext_ct_MSForm
                    '~~ A UserForm change required specific investigation in order to also detect a design change
                    CodeChanged = Me.FormCodeAndOrDesignChanged(ignore_empty_lines:=ignore_empty_lines)
            End Select
            fso.DeleteFile sTempExp
            If lKindOfComp = enClonedRaw Then Me.KindOfCodeChange = enKindOfCodeChange.enRawCloneOnly
        Case enRemoteRaw
            Select Case Me.VBComp.Type
                Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document
                    CodeChanged = _
                    mFl.sDiffer(dif_file1:=raw_clone.ExportFile _
                                , dif_file2:=raw_remote.ExportFile _
                                , dif_ignore_empty_records:=ignore_empty_lines _
                                 )
                Case vbext_ct_MSForm
                    '~~ A UserForm change required specific investigation in order to also detect a design change
                    CodeChanged = cComp.FormCodeAndOrDesignChanged(ignore_empty_lines:=ignore_empty_lines)
            End Select
    End Select

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
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
        Case mErH.ErrMsgDefaultButton: End
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
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Private Function DebugComp() As String
' ------------------------------------
'
' ------------------------------------
    
    Dim lMax        As Long
    
    If Not wb Is Nothing _
    Then lMax = lCompMaxLen _
    Else lMax = mDat.CommCompsMaxLenght()
    
    If lMax > 0 _
    Then DebugComp = sCompName & " " & String(lMax - Len(sCompName), ".") & ": " _
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
    
    With New WshShell
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

Private Function ExportFilePath() As String
' -----------------------------------------
' Returns the path for the code backup
' export file  of the Component which
' defaults to the Workbook's path with a
' Workbook's base name folder.
' Will be created if not existent.
' --------------------------------------
    
    Dim sPath As String

    sPath = sPathCommon & "\" & sWrkbkBaseName
    '~~ Make sure the returned export path specifies an existing folder
    With New FileSystemObject
        If Not .FolderExists(sPath) Then
            .CreateFolder (sPath)
        End If
    End With
    ExportFilePath = sPath

End Function

Public Function ExportFileToArray(ByVal fl As FILE) As String()
' ---------------------------------------------------------------------
' Returns the content of the file (vFile) which may be provided as file
' object or full file name string as array, considering any kind of
' line break characters.
' ---------------------------------------------------------------------
    Const PROC = "ExportFileToArray"
    Const ATTRIB_VB = "Attribute VB_"

    On Error GoTo eh
    Dim a           As Variant
    Dim i           As Long
    
    a = mFl.ToArray(fl)
    '~~ Remove the extra module information lines
    For i = UBound(a) To LBound(a) Step -1
        '~~ Find last Exported module info
        If Left(a(i), Len(ATTRIB_VB)) = ATTRIB_VB Then
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
    ExportFileToArray = a

xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Function ExportFileToDict(ByVal fl As FILE) As Dictionary
' --------------------------------------------------------------
' Returns the content of the export file (fl) as Dictionary.
' All lines preceeding "Attribute " lines and "Attribute " lines
' are excempted.
' --------------------------------------------------------------
    Const PROC = "ExportFileToArray"
    Const ATTRIB_LINE = "Attribute "

    On Error GoTo eh
    Dim i               As Long
    Dim dct             As Dictionary
    Dim v               As Variant
    Dim cllRemove       As New Collection
    Dim sLine           As String
    Dim cll             As Collection
    
    Set cll = New Collection
    Set dct = mFl.ToDict(fl)
    
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
    
    Set ExportFileToDict = dct

xt: Set cll = Nothing
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Private Function FileExists(ByVal sFile As String) As Boolean
    With New FileSystemObject:  FileExists = .FileExists(sFile):    End With
End Function

Public Function FormCodeAndOrDesignChanged( _
                            Optional ByVal ignore_empty_lines As Boolean = True) As Boolean
' ------------------------------------------------------------------------------------------
' Whether a UserForm's code and/or design has changed can only be checked by the comparison
' of the size of the current and a temporary ExportFile of the UserForm. When either the
' size and content of the .frm files are different or the size of the .frx files are
' different, a change is considered.
' ------------------------------------------------------------------------------------------
    Const PROC = "FormCodeAndOrDesignChanged"
    
    On Error GoTo eh
    Dim sTempPath   As String
    Dim sTempFolder As String
    Dim flTemp      As FILE
    
    With New FileSystemObject
        sTempFolder = .GetFile(flExport.PATH).ParentFolder
        sTempFolder = sTempFolder & "\Temp"
        If Not .FolderExists(sTempFolder) Then .CreateFolder sTempFolder
        
        sTempPath = sTempFolder & "\" & vbc.name & TypeExtention(vbc)
        vbc.Export sTempPath
        Set flTemp = mFl.GetFile(sTempPath)
        
        If UserFormSizeIsDifferent(sid_file_1:=Me.ExportFile, sid_file_2:=flTemp) Then
            FormCodeAndOrDesignChanged = True
        ElseIf mFl.sDiffer(dif_file1:=flExport _
                           , dif_file2:=flTemp _
                           , dif_ignore_empty_records:=ignore_empty_lines _
                            ) _
        Then
            FormCodeAndOrDesignChanged = True
        End If
        
        .GetFolder(sTempFolder).Delete
    End With
                                            
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Private Function GetFile(ByVal sFile As String) As FILE
    With New FileSystemObject:  Set GetFile = .GetFile(sFile):  End With
End Function

Public Sub RegisterAsHostedRaw()
    mDat.CommCompExpFileFullName(sCompName) = sExportFullName
    mDat.CommCompHostWorkbookBaseName(sCompName) = sWrkbkBaseName
End Sub

Public Sub RegisterAsHostWorkbook()
    mDat.CommCompsHostWorkbookFullName(sWrkbkBaseName) = sWrkbkFullName
End Sub

Public Sub ReplaceRemoteWithClonedRaw( _
                        ByVal raw_remote_comp_name As String, _
                        ByVal raw_remote_wrkbk_full_name As String, _
                        ByVal raw_cloned_exp_file As FILE)
' ---------------------------------------------------------
' Replace the Common Component's raw with the code of the
' used Common Component by means of its Export File.
' ---------------------------------------------------------
    Const PROC = "ReplaceRemoteWithClonedRaw"

    On Error GoTo eh
    Dim wbRemoteRaw     As Workbook
    Dim vbcRemoteRaw    As VBComponent
    Dim sTempName       As String: sTempName = raw_remote_comp_name & "_Temp"
    Dim i               As Long
    
    '~~ When the primary set temp name for the to-be-replace vbc is already existing
    '~~ a yet non-existing temp name has to be found because an existing vbc cannot
    '~~ be removed beforehand but only afterwards
    While mVBP.ComponentExists(vWb:=wb, vComp:=sTempName)
        i = i + 1: sTempName = sTempName & i
    Wend
    
    Set wbRemoteRaw = mWrkbk.GetOpen(cRemoteRaw.WrkbkFullName) ' Wrkbk is the Workbook hosting the raw Component
    Set vbcRemoteRaw = wbRemoteRaw.VBProject.VBComponents(Me.CompName)
    
    ' 1. Rename existing component
    vbcRemoteRaw.name = vbc.name & "_temp"
    ' 2. (Re-)Import the raw's clone Export File
    wbRemoteRaw.VBProject.VBComponents.Import cComp.ExportFileFullName
    ' 3. Remove any remaining temp vbc
    sTempName = Me.CompName & "_Temp"
    While mVBP.ComponentExists(vWb:=wbRemoteRaw, vComp:=sTempName)
        wbRemoteRaw.VBProject.VBComponents.Remove wb.VBProject.VBComponents(sTempName)
        i = i + 1
        sTempName = sTempName & i
    Wend
    Debug.Print DebugComp & Now() & " Remote raw Component replaced with Export File '" & cComp.ExportFileFullName & "'."

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub ReplaceClonedWithRemoteRaw(ByVal ru_vbc As VBComponent)
' ----------------------------------------------------------------
'
' ----------------------------------------------------------------
    Const PROC = "ReplaceClonedWithRemoteRaw"

    On Error GoTo eh
    Dim sTempName   As String: sTempName = ru_vbc.name & "_Temp"
    Dim i           As Long
    Dim fso         As New FileSystemObject
    
    While mVBP.ComponentExists(vWb:=wb, vComp:=sTempName)
        i = i + 1
        sTempName = sTempName & i
    Wend
    
    ru_vbc.name = sTempName
    With wb.VBProject.VBComponents
        .Import cRemoteRaw.ExportFileFullName
    End With
    Set vbc = wb.VBProject.VBComponents(sCompName) ' renew the vbc object to point to the new imported component
    
    sTempName = vbc.name & "_Temp"
    While mVBP.ComponentExists(vWb:=wb, vComp:=sTempName)
        wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(sTempName)
        i = i + 1
        sTempName = sTempName & i
    Wend
    
    On Error Resume Next
    Application.EnableEvents = False
    wb.Save
    Application.EnableEvents = True
    
    With Me ' The used class
        '~~ Log the updated code version
        .CodeVersionAsOfDate = cRemoteRaw.ExportFile.DateLastModified
        '~~ Log the update
        mCompManLog.LogAction log_wb:=Me.Wrkbk.name _
                            , log_action:=DebugComp & "The original's last update date = " & cRemoteRaw.ExportFile.DateLastModified
        mCompManLog.LogAction log_wb:=Me.Wrkbk.name _
                            , log_action:=DebugComp & "The registered last as-of used  = " & .CodeVersionAsOfDate
    End With

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function Space(ByVal l As Long) As String
' --------------------------------------------------
' Unifies the VB differences SPACE$ and Space$ which
' leads to code diferences where there aren't any.
' --------------------------------------------------
    Space = VBA.Space$(l)
End Function

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
    Dim wb      As Workbook
    Dim fso     As New FileSystemObject
                
    Debug.Print DebugComp & Now() & " About to synch the remote raw component's code with the code in the cloned raw's code in the Export File '" & cComp.ExportFileFullName & "'."
    '~~ Remove all lines from the target component
    Set vbcm = cRemoteRaw.VBComp.CodeModule
    With vbcm
        .DeleteLines 1, .CountOfLines
        '~~ Insert all code lines from the cloned raw component's Export File into the remote raws CodeModule
        iLine = 0
        For Each v In ExportFileToArray(cComp.ExportFile)
            iLine = iLine + 1
            vbcm.InsertLines iLine, v
        Next v
    End With
    
    Debug.Print DebugComp & Now() & " Code of remote raw component synchronized with the code in the cloned raw's Export File '" & cComp.ExportFileFullName & "'."
                
xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
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
        For Each v In ExportFileToArray(cRemoteRaw.ExportFile)
            iLine = iLine + 1
            .InsertLines iLine, v
        Next v
    End With
    
    Debug.Print DebugComp & Now() & " Code of the cloned raw component's CodeModule synchronized with code the remote raw component's Export File '" & cRemoteRaw.ExportFileFullName & "'."
                
xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Function TypeExtention(ByVal te_vbc As VBComponent) As String
    Select Case te_vbc.Type
        Case vbext_ct_StdModule:    TypeExtention = ".bas"
        Case vbext_ct_ClassModule:  TypeExtention = ".cls"
        Case vbext_ct_MSForm:       TypeExtention = ".frm"
        Case vbext_ct_Document:     TypeExtention = ".cls"
    End Select
End Function

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
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub ReplaceRemoteWithClonedRawWhenConfirmed( _
                                             ByRef rwu_updated As Boolean)
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
    Dim fso             As New FileSystemObject
    Dim wbRemoteRaw     As Workbook
    Dim vbcRemoteRaw    As VBComponent
    
    If ReplaceRemoteWithClonedRawConfirmed Then
            '~~ Get the origin host Workbook open
        With cRemoteRaw
            .WrkbkFullName = mDat.CommCompHostWorkbookFullName(cComp.CompName)
            .ExportFile = fso.GetFile(mDat.CommCompExpFileFullName(cComp.CompName))
            Set wbRemoteRaw = mWrkbk.GetOpen(.WrkbkFullName)
            Set vbcRemoteRaw = wbRemoteRaw.VBProject.VBComponents(.CompName)
            Select Case vbcRemoteRaw.Type
                Case vbext_ct_Document
                    '~~ A data module cannot be replaced like the other ones. Thus only its code
                    '~~ is synchronized - disregarding any consequences this may have when the
                    '~~ structure and design of the underlying Workbook or Worksheet has changed.
                    .SynchRemoteWithClonedRaw
                Case vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_MSForm
                    .ReplaceRemoteWithClonedRaw raw_remote_comp_name:=.CompName _
                                              , raw_remote_wrkbk_full_name:=.WrkbkFullName _
                                              , raw_cloned_exp_file:=.ExportFile
            End Select
            rwu_updated = True
        End With ' cRemoteRaw
    End If

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

'Public Sub UpdateUsedWithOrigin()
'' --------------------------------------------------
'' Update is done either by code synch or by replace.
'' --------------------------------------------------
'    Const PROC = "UpdateUsedWithOrigin"
'
'    On Error GoTo eh
'
'    With Me ' The used class
'        .ReplaceClonedWithRemoteRaw vbc
'
'        '~~ Log the updated code version
'        .CodeVersionAsOfDate = cRemoteRaw.ExportFile.DateLastModified
'        '~~ Log the update
'        mCompManLog.LogAction log_wb:=Me.Wrkbk.name _
'                            , log_action:=DebugComp & "The original's last update date = " & cRemoteRaw.ExportFile.DateLastModified
'        mCompManLog.LogAction log_wb:=Me.Wrkbk.name _
'                            , log_action:=DebugComp & "The registered last as-of used  = " & .CodeVersionAsOfDate
'    End With
'
'xt: Exit Sub
'
'eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
'        Case mErH.DebugOpt1ResumeError: Stop: Resume
'        Case mErH.DebugOpt2ResumeNext: Resume Next
'        Case mErH.ErrMsgDefaultButton: End
'    End Select
'End Sub

Private Function UserFormSizeIsDifferent( _
                                   ByVal sid_file_1 As FILE, _
                                   ByVal sid_file_2 As FILE) As Boolean
    Dim sFrxFile1Path  As String
    Dim sFrxFile2Path  As String

    sFrxFile1Path = Replace$(sid_file_1.PATH, ".frm", ".frx")
    sFrxFile2Path = Replace$(sid_file_2.PATH, ".frm", ".frx")
    
    With New FileSystemObject
        If sid_file_1.Size <> sid_file_2.Size Then
            UserFormSizeIsDifferent = True
        ElseIf .GetFile(sFrxFile1Path).Size <> .GetFile(sFrxFile2Path).Size Then
            UserFormSizeIsDifferent = True
        End If
    End With
    
End Function

