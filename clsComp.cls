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
'       Represents a VBComponent with extended (Management) properties and methods.
'       The Component may be of either of the following kinds:
'       - a used Common Component
'       - a hosted Common Component, i.e. the origin code of a Common Component
'       - neither of the above, i.e. a "normal" VBProject Component
'
' Properties:
' - ExportFile              R/W File object created by menas of the BackUpCode method
' - HostedCommonComponents  W   String with named Components hosted in the Workbook
' - IsUsedCommonComponent                R   TRUE when a Component is regarded a Common one
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
Private cCommonOrigin   As clsComp      ' refers to the code source for a kind of "used" Common Component
Private sHostBaseName   As String       ' Common Component host Workbook's base name
Private sHostFullName   As String       ' Common Component host Workbook's full name
Private sComponentName  As String
Private wbHost          As Workbook     ' Common Component host Workbook
Private vbc             As VBComponent  ' VBComponent managed by this Class Module
'Private vbcKind         As enKindOfComp
Private dtUpdatedAsOf   As Date
Private flExport        As File
Private foExport        As Folder
Private dctCommonHosted As Dictionary
Private sPathCommon     As String
Private sExportFullName As String       ' Full filespec for the Export File created by the method BackUpCode
Private lCompMaxLen     As Long
Private dctCommComps    As Dictionary

Private Sub Class_Initialize()
    sPathCommon = mCommDat.CommonComponentsBasePath
    lCompMaxLen = mCommDat.CommCompsMaxLenght
    Set dctCommonHosted = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set cCommonOrigin = Nothing
End Sub

Public Property Get CodeVersionAsOfDate() As Date
    CodeVersionAsOfDate = mCommDat.CodeVersionAsOfDate(sHostBaseName, sComponentName)
End Property
Public Property Let CodeVersionAsOfDate(ByVal dtAsOf As Date):
    mCommDat.CodeVersionAsOfDate(sHostBaseName, sComponentName) = dtAsOf
End Property
Public Property Let ComponentName(ByVal s As String):   sComponentName = s:             End Property
Public Property Get ComponentName() As String:          ComponentName = sComponentName: End Property

Public Property Get CodeChanged( _
                 Optional ByVal ignore_empty_lines As Boolean = False) As Boolean
' -------------------------------------------------------------------------------
' Returns TRUE
' - when the temporary export file of the Component compared with its last Export
'   File indicates a  change
' - when the Export File does not exist.
' For the compare the Component is temporary exported !
' -------------------------------------------------------------------------------
    Const PROC  As String = "CodeChanged"
       
    On Error GoTo eh
    Dim dctComp As Dictionary
    Dim dctFile As Dictionary
    Dim fso     As New FileSystemObject
    Dim sTmpExp As String
    Dim a1()    As String
    Dim a2()    As String
    
    sTmpExp = Replace(Me.ExportFileFullName, Me.Extension, "_Temp" & Me.Extension)
    With fso
        If Not .FileExists(Me.ExportFileFullName) Then
            CodeChanged = True
            GoTo xt
        End If
    End With
        
    Select Case vbc.Type
        Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document
        
            If vbc.Name <> BaseName(flExport.Path) _
            Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The to be compared sources do not have the same name!"
            
            Me.VBComp.Export sTmpExp ' any existing would be overwritten
            CodeChanged = _
            mFile.sDiffer(dif_file1:=fso.GetFile(sTmpExp) _
                        , dif_file2:=flExport _
                        , dif_ignore_empty_records:=ignore_empty_lines _
                         )
        
        Case vbext_ct_MSForm
            '~~ A UserForm change required specific investigation in order to also detect a design change
            CodeChanged = FormCodeAndOrDesignChanged(ignore_empty_lines:=ignore_empty_lines)
    End Select

xt: If fso.FileExists(sTmpExp) Then fso.DeleteFile sTmpExp ' cleanup
    Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Private Property Get CommonOriginCodeDiffersFromCommonUsedCode( _
                    ByVal flOrigin As File, _
                    ByVal flUsed As File) As Boolean
' ----------------------------------------------------------------
' Returns TRUE when the used Common Component's Export File
' differs from the origin Common Component's Export File.
' Note: Code comparison by means of the Export files allows to
'       check the code change with a usually closed host Workbook.
' ----------------------------------------------------------------
    Const PROC = "CommonOriginCodeDiffersFromCommonUsedCode"
    
    On Error GoTo eh
    Dim aFile1  As Variant
    Dim aFile2  As Variant
    
    aFile1 = ExportFileToArray(flOrigin) ' obtained through its name with mCommDat
    aFile2 = ExportFileToArray(flUsed)
    CommonOriginCodeDiffersFromCommonUsedCode = mBasic.ArrayDiffers(aFile1, aFile2)

xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Public Property Get CommonOriginCodeChanged( _
                    Optional ByRef flCommonOriginExportFile As File, _
                    Optional ByRef sCommonOriginHostFullName As String) As Boolean
' --------------------------------------------------------------------------------
' Returns TRUE when the common origin code's ExportFile date is no longer equal to
' the used code's as-of-date registered with the last update.
' --------------------------------------------------------------------------------
    With Me
        If .CodeChanged(ignore_empty_lines:=True) Then
            '~~ Any difference constituted only by empty code lines is ignored!
            .BackUpCode ' Just in case the used and changed code has yet not been exported
        End If
        If .IsCommonUsed(flCommonOriginExportFile, sCommonOriginHostFullName) Then
            If flCommonOriginExportFile.DateLastModified <> .CodeVersionAsOfDate Then
                If CommonOriginCodeDiffersFromCommonUsedCode(.CommonOrigin.ExportFile, .ExportFile) Then
                    CommonOriginCodeChanged = True
                End If
            End If
        End If
    End With

End Property

Public Property Get ExportFile() As File
    
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

Public Property Let ExportFile(ByVal fl As File):  Set flExport = fl:  End Property

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

Public Property Let HostedCommonComponents( _
                    ByVal sHosted As String)
' -------------------------------------------
' Saves the name of hoste Common Components
' (sHosted) to the Dictionary (dctCommonHosted).
' -------------------------------------------
Dim v       As Variant
Dim sComp   As String

    Set dctCommonHosted = New Dictionary
    For Each v In Split(sHosted, ",")
        sComp = Trim$(v)
        dctCommonHosted.Add sComp, sComp
    Next v
    If dctCommonHosted.Count <> 0 Then
        mCommDat.CommCompsHostWorkbookFullName(BaseName(wbHost.Name)) = wbHost.FullName
    End If
    
End Property

Public Property Get IsCommonUsed( _
                    ByRef fl As File, _
                    ByRef s As String) As Boolean
' -----------------------------------------------
' IsCommonUsed is the oposite of IsHostedCommonComponent.
' When the Component is a Common used one the
' Export File object is returned.
' -----------------------------------------------
    If Me.IsUsedCommonComponent Then
        If Not IsHostedCommonComponent Then
            IsCommonUsed = True
            mCommDat.IsCommonComponent sComponentName, fl, s
        End If
    End If

End Property

Public Property Get KindOfChange( _
       Optional ByVal dtSource As Date, _
       Optional ByVal dtTarget As Date) As enKindOfChange
' -------------------------------------------------------
' Returns the kind of code change based on the comparison
' of the LastModified date of the source and a target
' Export File.
' -------------------------------------------------------
    
    Dim dtAsOf      As Date
    Dim fl          As File

    With Me.CommonOrigin
        .Name = sComponentName
        With New FileSystemObject
            Set fl = .GetFile(mCommDat.CommCompExpFileFullName(sComponentName))
        End With
        Set flExport = fl
    End With
    
    If IsMissing(dtSource) Then dtSource = cCommonOrigin.ExportFile.DateLastModified
    If IsMissing(dtTarget) Then dtTarget = flExport.DateLastModified
    
    dtAsOf = Me.UpdatedAsOf
    
    If dtSource > dtAsOf Then
        '~~ The origin code source of the Common Component had chaged
        If dtTarget = dtSource Then
            '~~ Target has not changed
            KindOfChange = enKindOfChange.enSourceOnly
        ElseIf dtTarget > dtSource Then
            '~~ Target has changed
            KindOfChange = enKindOfChange.enSourceAndTarget
        End If
    ElseIf dtSource = dtAsOf Then
        '~~ Source had not changed
        If dtTarget <= dtSource Then
            '~~ Target had not changed
            Stop ' None had changed ?!?!
            KindOfChange = enKindOfChange.enNeitherNore
        ElseIf dtTarget > dtSource Then
            '~~ Target has changed
            KindOfChange = enKindOfChange.enTargetOnly
        End If
    End If
    
End Property

Public Property Get Name() As String:                       Name = sComponentName:                               End Property

Public Property Let Name(ByVal s As String):                sComponentName = s:                                  End Property

Public Property Get CommonOrigin() As clsComp
    If cCommonOrigin Is Nothing Then Set cCommonOrigin = New clsComp
    Set CommonOrigin = cCommonOrigin
End Property

Public Property Let CommonOrigin(ByVal c As clsComp):       Set cCommonOrigin = c:                      End Property

Public Property Get UpdatedAsOf() As Date:                  UpdatedAsOf = dtUpdatedAsOf:                End Property

Public Property Let UpdatedAsOf(ByVal dt As Date):          dtUpdatedAsOf = dt:                         End Property

Private Property Get UpdateOriginWithUsedConfirmed() As enUpdateReply
' -------------------------------------------------------------------
' Returns TRUE when the "reverse" code update is confirmed.
' -------------------------------------------------------------------
    Const PROC                  As String = "UpdateOriginWithUsedConfirmed"
    Const REPLY_UPDATE_ORIGIN   As String = "Update Origin Code"
    Const REPLY_DISCARD_CHANGES As String = "Discard all Code Changes"
    Const REPLY_DSPLY_DIFF      As String = "Show Difference"
    Const REPLY_ABORT           As String = "Abort"
    
    Dim sReplyDsplyDiff         As String
    Dim sBaseMsg                As String
    Dim sMsg                    As String
    Dim sDsplyDiffMsg           As String
    
    On Error GoTo eh
    
    '~~ Prepare communication message and reply option
    If mBasic.AppIsInstalled("WinMerge") Then
        sReplyDsplyDiff = "," & REPLY_DSPLY_DIFF
        sDsplyDiffMsg = DQUOTE & REPLY_DSPLY_DIFF & DQUOTE & " to see what has been changed and decide then (leave with Esc key)"
    Else
        sReplyDsplyDiff = vbNullString
        sDsplyDiffMsg = "."
    End If
                
    sBaseMsg = "The Common Component '" & sComponentName & "' used in this Workbook has been changed in it." & vbLf & vbLf & _
               "Reply with:" & vbLf & _
               DQUOTE & REPLY_UPDATE_ORIGIN & DQUOTE & " when the changes should become common/permanent" & vbLf & _
               sDsplyDiffMsg & vbLf & _
               DQUOTE & REPLY_DISCARD_CHANGES & DQUOTE & " to discard all changes, i.e. (re-)instate the origin code"
    
    '~~ Offer options
    sMsg = sBaseMsg
    sDsplyDiffMsg = sBaseMsg & sDsplyDiffMsg
    Select Case mMsg.Box(dsply_title:="Common Component code changed in this Workbook", _
                         dsply_msg:=sDsplyDiffMsg, _
                         dsply_buttons:=mMsg.Buttons(REPLY_UPDATE_ORIGIN & sReplyDsplyDiff, REPLY_DISCARD_CHANGES, REPLY_ABORT) _
                        )
        Case REPLY_UPDATE_ORIGIN    ' Update without display of the differences
            UpdateOriginWithUsedConfirmed = True
            Exit Property
        
        Case REPLY_DSPLY_DIFF
            DisplayDiff flLeft:=Me.CommonOrigin.ExportFile, _
                        flRght:=flExport, _
                        sTitleLeft:="Code of " & sComponentName & " currently hosted in Workbook/VBProject " & Me.CommonOrigin.HostBaseName, _
                        sTitleRght:="Code of " & sComponentName & " currently used in Workbook/VBProject " & Me.HostBaseName
            '~~ When the code difference display has been closed, offer update/ignore again
            Select Case mMsg.Box(dsply_title:="Common Component code changed in this Workbook", _
                                 dsply_msg:=sMsg, _
                                 dsply_buttons:=mMsg.Buttons(REPLY_UPDATE_ORIGIN, REPLY_DISCARD_CHANGES, REPLY_ABORT) _
                                )
                Case REPLY_UPDATE_ORIGIN:   UpdateOriginWithUsedConfirmed = enUpdateOriginWithUsed
                Case REPLY_DISCARD_CHANGES: UpdateOriginWithUsedConfirmed = enUpdateUsedWithOrigin
                Case REPLY_ABORT:           UpdateOriginWithUsedConfirmed = enUpdateNone
            End Select
         
        Case REPLY_UPDATE_ORIGIN:   UpdateOriginWithUsedConfirmed = enUpdateOriginWithUsed
        Case REPLY_DISCARD_CHANGES: UpdateOriginWithUsedConfirmed = enUpdateUsedWithOrigin
        Case REPLY_ABORT:           UpdateOriginWithUsedConfirmed = enUpdateNone
        
    End Select
                                            
xt: Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Public Property Get VBComp() As VBComponent
' -----------------------------------------
' When the class object represents a source
' Component there is no VBComp object (vbc)
' and thus the Component object is obtained
' from the source Workbook known by its
' FullName.
' -----------------------------------------
    If vbc Is Nothing _
    Then Set VBComp = Me.CommonOrigin.Host.VBProject.VBComponents(sComponentName) _
    Else Set VBComp = vbc
End Property

Public Property Let VBComp(ByVal oComp As VBComponent)
    Const PROC = "VBComp_Let"

    On Error GoTo eh
    If Me.Host Is Nothing _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "There's yet no Host Workbook assigned!"
    
    Set vbc = oComp
    sComponentName = vbc.Name
    sExportFullName = ExportFilePath & "\" & sComponentName & ExportFileExtention
    If FileExists(sExportFullName) Then
        Set flExport = GetFile(sExportFullName)
    Else
        '~~ This is the very first time the component comes to management
        Me.BackUpCode
    End If
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Property

Public Property Get HostBaseName() As String:                 HostBaseName = BaseName(sHostFullName):         End Property

Public Property Get HostFullName() As String:                 HostFullName = sHostFullName:                   End Property

Public Property Let HostFullName(ByVal s As String):          sHostFullName = s:                            End Property

Public Property Get Host() As Workbook
    If wbHost Is Nothing _
    Then Set wbHost = mWrkbk.GetOpen(sHostFullName) _
    Else Set Host = wbHost
End Property

Public Property Let Host(ByVal wbk As Workbook)
    
    Dim fso As New FileSystemObject
    
    Set wbHost = wbk
    sHostFullName = wbHost.FullName
    sHostBaseName = fso.GetBaseName(sHostFullName)
    Set foExport = fso.GetFolder(ExportFilePath)
    For Each vbc In wbHost.VBProject.VBComponents
        lCompMaxLen = mBasic.Max(lCompMaxLen, Len(vbc.Name))
    Next vbc

End Property

Public Sub BackUpCode()
' -------------------------------------
' Exports the VBComponent (vbc)
' provided the CodeModule is not empty.
' -------------------------------------
    
    If Not vbc.CodeModule.CountOfLines = 0 Then
        If Me.CodeChanged Then
            vbc.Export sExportFullName
        End If
        Set flExport = GetFile(sExportFullName)
    End If
    mCompLog.LogAction wbHost.Name, sLog:=DebugComp
    mCompLog.LogAction wbHost.Name, sLog:=DebugComp & "Code has changed, code-backup (ExportFile) re-newed"

End Sub

Private Function ExportFilePath() As String
' -----------------------------------------
' Returns the path for the code backup
' export file  of the Component which
' defaults to the Workbook's path with a
' Workbook's base name folder.
' Will be created if not existent.
' --------------------------------------
    
    Dim sPath As String

    sPath = sPathCommon & "\" & sHostBaseName
    '~~ Make sure the returned export path specifies an existing folder
    With New FileSystemObject
        If Not .FolderExists(sPath) Then
            .CreateFolder (sPath)
        End If
    End With
    ExportFilePath = sPath

End Function

Private Function FormCodeAndOrDesignChanged( _
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
    Dim flTemp      As File
    
    With New FileSystemObject
        sTempFolder = .GetFile(flExport.Path).ParentFolder
        sTempFolder = sTempFolder & "\Temp"
        If Not .FolderExists(sTempFolder) Then .CreateFolder sTempFolder
        
        sTempPath = sTempFolder & "\" & vbc.Name & TypeExtention(vbc)
        vbc.Export sTempPath
        Set flTemp = mFile.GetFile(sTempPath)
        
        If UserFormSizeIsDifferent(Me.ExportFile, flTemp) Then
            FormCodeAndOrDesignChanged = True
        ElseIf mFile.sDiffer(dif_file1:=flExport _
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

Public Function IsUsedCommonComponent() As Boolean
' -------------------------------------------------------------
' Returns:
' - TRUE when the Component (sComponentName) is a known Common Component
' - When TRUE, the name of the Export File as file object (fl)
'   and the full name of the host Workbook (s).
' -------------------------------------------------------------
    If dctCommComps Is Nothing Then Set dctCommComps = mCommDat.KnownCommonComponents
    IsUsedCommonComponent = dctCommComps.Exists(sComponentName)
End Function

Public Property Get IsHostedCommonComponent() As Boolean
    IsHostedCommonComponent = dctCommonHosted.Exists(sComponentName)
End Property

Public Sub LogCommonCompExport()
    mCompLog.LogAction wbHost.Name, sComponentName & mBasic.Space$(lCompMaxLen - Len(vbc.Name)) & ": The original's last update date = " & flExport.DateLastModified
    mCompLog.LogAction wbHost.Name, sComponentName & mBasic.Space$(lCompMaxLen - Len(vbc.Name)) & ": The registered last as-of used  = " & Me.CodeVersionAsOfDate
End Sub

Public Sub UpdateOriginWithUsedWhenConfirmedByUser(ByRef bUpdated As Boolean)
' ---------------------------------------------------------------------------
' Updates the Common Component's origin code with the used Common Component's
' Export File provided the user confirmed this "reverse" update. When updated,
' TRUE is returned (bUpdated).
' !! Update is done either by code synch or by replace.
' ---------------------------------------------------------------------------
    Const PROC = "UpdateOriginWithUsedWhenConfirmedByUser"
    
    On Error GoTo eh
    Dim fl      As File
    Dim sWbHost As String
    
    Set fl = Me.ExportFile
    With Me.CommonOrigin
        Select Case UpdateOriginWithUsedConfirmed
            Case enUpdateOriginWithUsed
                '~~ Get the origin host Workbook open
                sWbHost = mCommDat.CommCompHostWorkbookFullName(sComponentName)
                .Host = mWrkbk.GetOpen(sWbHost)
                .VBComp = .Host.VBProject.VBComponents(sComponentName)
                Select Case .VBComp.Type
                    Case vbext_ct_ClassModule, vbext_ct_Document, vbext_ct_StdModule
                        .SynchOriginWithUsed
                    Case vbext_ct_MSForm
                        .ReplaceOriginWithUsed
                End Select
                bUpdated = True
            
            Case enUpdateUsedWithOrigin
                '~~ Discard changes by updating the used code with the origin code
                Debug.Print "Re-establish the origin code from its export/backup file"
                Select Case .VBComp.Type
                    Case vbext_ct_ClassModule, vbext_ct_Document, vbext_ct_StdModule
                        .SynchUsedWithOrigin
                    Case vbext_ct_MSForm
                        .ReplaceUsedWithOrigin
                End Select
            
            Case enUpdateNone
        End Select
    End With ' Me.CommonOrigin

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub UpdateUsedWithOrigin()
' -------------------------------
' Update is done either by code
' synch or by replace.
' -------------------------------
Const PROC  As String = "UpdateUsedWithOrigin"

    On Error GoTo eh
    
    With Me
        .ReplaceUsedWithOrigin
        
        '~~ Log the updated code version
        .CodeVersionAsOfDate = .CommonOrigin.ExportFile.DateLastModified
        '~~ Log the update
        mCompLog.LogAction wbHost.Name, sComponentName & mBasic.Space$(lCompMaxLen - Len(sComponentName)) & ": The original's last update date = " & .CommonOrigin.ExportFile.DateLastModified
        mCompLog.LogAction wbHost.Name, sComponentName & mBasic.Space$(lCompMaxLen - Len(sComponentName)) & ": The registered last as-of used  = " & .CodeVersionAsOfDate
    End With

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub SynchUsedWithOrigin()
' ------------------------------
' Synchronizes used Common
' Component's code with the
' origin's Export File.
' ------------------------------
    Const PROC = "SynchUsedWithOrigin"

    On Error GoTo eh
    Dim iLine   As Long
    Dim v       As Variant
    Dim vbcm    As CodeModule
                
    '~~ Remove all lines from the target component
    Set vbcm = vbc.CodeModule
    With vbcm
        .DeleteLines 1, .CountOfLines
    End With
    
    '~~ Insert/copy all code lines from the source to the target component
    iLine = 0
    For Each v In ExportFileToArray(Me.CommonOrigin.ExportFile)
        iLine = iLine + 1
        vbcm.InsertLines iLine, v
    Next v
    Debug.Print DebugComp & Now() & " Code in " & Me.Host.Name & " synchronized with code in " & Me.CommonOrigin.Host.Name & " ( via Export File '" & Me.CommonOrigin.ExportFile.Name & "'."
                
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub SynchOriginWithUsed()
' ------------------------------
' Synchronizes the Common
' Component's origin code with
' the Export File of the used
' Common Component.
' ------------------------------
    Const PROC = "SynchOriginWithUsed"

    On Error GoTo eh
    Dim iLine   As Long
    Dim v       As Variant
    Dim vbcm    As CodeModule
    Dim wbHost  As Workbook
                
    Debug.Print DebugComp & Now() & " About to synch code with Export File '" & Me.CommonOrigin.ExportFile.Name & "'."
    '~~ Remove all lines from the target component
    Set wbHost = mWrkbk.GetOpen(Me.CommonOrigin.HostFullName)
    Set vbcm = wbHost.VBProject.VBComponents(sComponentName).CodeModule
    With vbcm
        .DeleteLines 1, .CountOfLines
    End With
    
    '~~ Insert/copy all code lines from the source to the target component
    iLine = 0
    For Each v In ExportFileToArray(Me.CommonOrigin.ExportFile)
        iLine = iLine + 1
        vbcm.InsertLines iLine, v
    Next v
    
    Debug.Print DebugComp & Now() & " Code synchronized with Export File '" & Me.CommonOrigin.ExportFile.Name & "'."
                
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub ReplaceUsedWithOrigin()
    Const PROC = "ReplaceUsedWithOrigin"

    On Error GoTo eh
    
    CleanExportFile Me.ExportFile
    vbc.Name = vbc.Name & "_Temp"
    With wbHost.VBProject.VBComponents
        .Import Me.CommonOrigin.ExportFile.Path
        .Remove vbc
    End With
    Set vbc = wbHost.VBProject.VBComponents(sComponentName)
    On Error Resume Next
    Debug.Print DebugComp & Now() & " Replaced with Export File '" & Me.CommonOrigin.ExportFile.Name & "' from Workbook " & Me.CommonOrigin.Host.Name
    If Err.Number <> 0 Then Debug.Print DebugComp & Now() & " Replaced with Export File '" & Me.CommonOrigin.ExportFile.Name & "' from Workbook " & Me.CommonOrigin.Host.Name
    Application.EnableEvents = False
    wbHost.Save
    Application.EnableEvents = True
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub ReplaceOriginWithUsed()
' --------------------------------
' Replace the Common Component's
' origin code with the used ones
' Export File.
' --------------------------------
    Const PROC = "ReplaceOriginWithUsed"

    On Error GoTo eh
    Dim wbHost      As Workbook
    
    Debug.Print DebugComp & Now() & " About to replace with Export File '" & Me.CommonOrigin.ExportFile.Name & "'."
    CleanExportFile Me.ExportFile
    With Me.CommonOrigin
        Set wbHost = mWrkbk.GetOpen(.HostFullName)
        wbHost.Remove vbc
        wbHost.Import Me.ExportFile.Path
        .VBComp = wbHost.VBProject.VBComponents(sComponentName)
    End With
    Debug.Print DebugComp & Now() & " Replaced with Export File '" & Me.CommonOrigin.ExportFile.Name & "'."

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function FileExists(ByVal sFile As String) As Boolean
    With New FileSystemObject:  FileExists = .FileExists(sFile):    End With
End Function

Private Function GetFile(ByVal sFile As String) As File
    With New FileSystemObject:  Set GetFile = .GetFile(sFile):  End With
End Function

Private Function BaseName(ByVal s As String) As String
    With New FileSystemObject:  BaseName = .GetBaseName(s): End With
End Function

Public Function TypeExtention(ByVal vbc As VBComponent) As String
    Select Case vbc.Type
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
    
eh: mErH.ErrMsg ErrSrc(PROC)
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComp" & "." & sProc
End Function

Public Function ExportFileToArray(ByVal fl As File) As String()
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
    ExportFileToArray = a

xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Function ExportFileToDict(ByVal fl As File) As Dictionary
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
    
    Set ExportFileToDict = dct

xt: Set cll = Nothing
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub DisplayDiff(ByVal flLeft As File, _
                        ByVal flRght As File, _
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
               """" & flLeft.Path & """" & " " & _
               """" & flRght.Path & """"
    
    With New WshShell
        .Run sCommand, windowStyle, waitOnReturn
    End With
    
End Sub

Private Function DebugComp() As String
' ------------------------------------
'
' ------------------------------------
    
    Dim lMax As Long

    If Not wbHost Is Nothing _
    Then lMax = lCompMaxLen _
    Else lMax = mCommDat.CommCompsMaxLenght
    If lMax > 0 _
    Then DebugComp = sComponentName & " " & String(lMax - Len(sComponentName), ".") & ": " _
    Else DebugComp = sComponentName & " : "

End Function

Private Function UserFormSizeIsDifferent(ByVal f1 As File, _
                                         ByVal f2 As File) As Boolean
    Dim sFrxFile1Path  As String
    Dim sFrxFile2Path  As String

    sFrxFile1Path = Replace$(f1.Path, ".frm", ".frx")
    sFrxFile2Path = Replace$(f2.Path, ".frm", ".frx")
    
    With New FileSystemObject
        If f1.Size <> f2.Size Then
            UserFormSizeIsDifferent = True
        ElseIf .GetFile(sFrxFile1Path).Size <> .GetFile(sFrxFile2Path).Size Then
            UserFormSizeIsDifferent = True
        End If
    End With
    
End Function

Public Sub RegisterAsHostWorkbook()
    mCommDat.CommCompsHostWorkbookFullName(sHostBaseName) = sHostFullName
End Sub

Public Sub RegisterAsCommonComponent()
    mCommDat.CommCompExpFileFullName(sComponentName) = sExportFullName
    mCommDat.CommCompHostWorkbookBaseName(sComponentName) = sHostBaseName
End Sub

Public Sub CleanExportFile(ByRef fl As File)
' ------------------------------------------
' Returns a clean export file, i.e. an empty
' line after the Attribute section removed.
' ------------------------------------------
    Const ATTRIBUTE_ID = "Attribute "
    
    Dim a   As Variant
    Dim i   As Long
    Dim ts  As TextStream

    a = mFile.ToArray(fl.Path)
    Do Until left(a(i), Len(ATTRIBUTE_ID)) <> ATTRIBUTE_ID
        i = i + 1
        If i > UBound(a) Then Exit Do
    Loop
    '~~ End of "Attribute " section reached
    If Len(a(i)) < 2 Then
        '~~ The "Atribute " section ends with an empty line
        mBasic.ArrayRemoveItems a, Index:=i
    End If

    With New FileSystemObject
        Set ts = .OpenTextFile(fl.Path, ForWriting, True, TristateFalse)
    End With
    For i = LBound(a) To UBound(a)
        If i = UBound(a) And Len(a(i)) <= 2 Then Exit For
        ts.WriteLine a(i)
    Next i
    ts.Close

End Sub

Public Function Exists(ByVal sComp As String) As Boolean
Dim vbc As VBComponent

    On Error Resume Next
    Set vbc = Me.Host.VBProject.VBComponents(sComp)
    Exists = Err.Number = 0
    Set vbc = Nothing
End Function

Public Function SelectedComponents(ByRef bWhenChangedOnly As Boolean) As Dictionary

    With fSelect
        .Host = Me.Host
        .show
        Set SelectedComponents = .SelectedComps
        bWhenChangedOnly = .WhenChangedOnly
    End With
    Unload fSelect

End Function

