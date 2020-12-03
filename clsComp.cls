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
Private cOrigin         As clsComp      ' refers to the code source for a kind of "used" Common Component
Private sHostBaseName   As String       ' Common Component host Workbook's base name
Private sHostFullName   As String       ' Common Component host Workbook's full name
Private sComponentName  As String       ' The class module's global component name
Private wb              As Workbook     ' The class' Workbook
Private vbc             As VBComponent  ' VBComponent managed by this Class Module
Private dtUpdatedAsOf   As Date
Private flExport        As File
Private foExport        As Folder
Private dctCommonHosted As Dictionary
Private sPathCommon     As String
Private sExportFullName As String       ' Full filespec for the Export File created by the method BackUpCode
Private lCompMaxLen     As Long
Private dctCommComps    As Dictionary

Private Sub Class_Initialize()
    sPathCommon = mCompManCfg.CommonComponentsBasePath
    lCompMaxLen = mCommDat.CommCompsMaxLenght
    Set dctCommonHosted = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set cOrigin = Nothing
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
    Dim fso     As New FileSystemObject
    Dim sTmpExp As String
    
    sTmpExp = Replace(Me.ExportFileFullName, Me.Extension, "_Temp" & Me.Extension)
    With fso
        If Not .FileExists(Me.ExportFileFullName) Then
            CodeChanged = True
            GoTo xt
        End If
    End With
        
    Select Case vbc.Type
        Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document
        
            If vbc.name <> BaseName(flExport.Path) _
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

Public Property Get OriginCodeHasChanged() As Boolean
' -------------------------------------------------------------------------------
' Returns TRUE when the common origin code's ExportFile date is no longer equal
' to the used code's as-of-date registered with the last update. Only performed
' for used component which is regarded Common, i.e. the origin code comes from
' a Workbook which "hosts" the component.
' -------------------------------------------------------------------------------
    With Me
        If .CodeChanged() Then .BackUpCode ' just in case ...
        If .ExportFile.DateLastModified <> .CompOrigin.ExportFile.DateLastModified Then
            '~~ Empty code lines do not constitute a relevant code change even when the used
            '~~ component has been backed up - which is also considering empty code lines
            If mFile.sDiffer(dif_file1:=.CompOrigin.ExportFile _
                           , dif_file2:=.ExportFile _
                           , dif_ignore_empty_records:=True) _
            Then OriginCodeHasChanged = True
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
        mCommDat.CommCompsHostWorkbookFullName(BaseName(wb.name)) = wb.FullName
    End If
    
End Property

Public Property Get IsCommonUsed( _
                    ByRef fl As File, _
                    ByRef s As String) As Boolean
' -----------------------------------------------
' IsCommonUsed is the oposite of IsHostedCommon.
' When the Component is a Common used one the
' Export File object is returned.
' -----------------------------------------------
    If Me.IsUsedCommonComponent Then
        If Not IsHostedCommon Then
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

    With Me.CompOrigin
        .name = sComponentName
        With New FileSystemObject
            Set fl = .GetFile(mCommDat.CommCompExpFileFullName(sComponentName))
        End With
        Set flExport = fl
    End With
    
    If IsMissing(dtSource) Then dtSource = cOrigin.ExportFile.DateLastModified
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

Public Property Get name() As String:                       name = sComponentName:                      End Property

Public Property Let name(ByVal s As String):                sComponentName = s:                         End Property

Public Property Get CompOrigin() As clsComp
    If cOrigin Is Nothing Then Set cOrigin = New clsComp
    Set CompOrigin = cOrigin
End Property

Public Property Let CompOrigin(ByVal c As clsComp):         Set cOrigin = c:                            End Property

Public Property Get UpdatedAsOf() As Date:                  UpdatedAsOf = dtUpdatedAsOf:                End Property

Public Property Let UpdatedAsOf(ByVal dt As Date):          dtUpdatedAsOf = dt:                         End Property

Private Property Get UpdateOriginWithUsedConfirmed() As Boolean
' -------------------------------------------------------------
' Returns TRUE when the "reverse" code update is confirmed.
' --------------------------------------------------------------
    Const PROC = "UpdateOriginWithUsedConfirmed"
    
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
    sTitle = "A ""Common Component's"" code has been changed in this Workbook"
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
    With sMsg
        .section(1).sLabel = "About this Component:"
        .section(1).sText = "The module is regarded ""Common"" because it is indacted ""hosted"" in another Workbook."
        .section(2).sLabel = "About this code change:"
        .section(2).sText = "The module, updated when this Workbook has opened, is no longer identical with its origin. " & _
                            "I.e. the component's code has been changed in this Workbook using it instead of in the " & _
                            "Workbook the origin is ""hosted""."
        .section(3).sText = "Take your decission or postpone it after having displayed the difference."
    End With
        
    Do
        sReply = mMsg.Dsply(dsply_title:=sTitle, _
                            dsply_msg:=sMsg, _
                            dsply_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sReplyUpdateOrigin
                UpdateOriginWithUsedConfirmed = True
                Exit Do
        
            Case sReplyDsplyDiff
                DisplayDiff flLeft:=Me.CompOrigin.ExportFile, _
                            flRght:=flExport, _
                            sTitleLeft:="Code of " & sComponentName & " currently hosted in Workbook/VBProject " & Me.CompOrigin.HostBaseName, _
                            sTitleRght:="Code of " & sComponentName & " currently used in Workbook/VBProject " & Me.HostBaseName
                ' The question is re-displayed for a decison
            Case sReplyIgnore
                UpdateOriginWithUsedConfirmed = False
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
' -----------------------------------------
' When the class object represents a source
' Component there is no VBComp object (vbc)
' and thus the Component object is obtained
' from the source Workbook known by its
' FullName.
' -----------------------------------------
    If vbc Is Nothing _
    Then Set VBComp = Me.CompOrigin.Wrkbk.VBProject.VBComponents(sComponentName) _
    Else Set VBComp = vbc
End Property

Public Property Let VBComp(ByVal oComp As VBComponent)
    Const PROC = "VBComp_Let"

    On Error GoTo eh
    If Me.Wrkbk Is Nothing _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "There's yet no Wrkbk Workbook assigned!"
    
    Set vbc = oComp
    sComponentName = vbc.name
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

Public Property Get Wrkbk() As Workbook
    If wb Is Nothing _
    Then Set wb = mWrkbk.GetOpen(sHostFullName) _
    Else Set Wrkbk = wb
End Property

Public Property Let Wrkbk(ByVal wbk As Workbook)
    
    Dim fso As New FileSystemObject
    
    Set wb = wbk
    sHostFullName = wb.FullName
    sHostBaseName = fso.GetBaseName(sHostFullName)
    Set foExport = fso.GetFolder(ExportFilePath)
    For Each vbc In wb.VBProject.VBComponents
        lCompMaxLen = mBasic.Max(lCompMaxLen, Len(vbc.name))
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
            mCompManLog.LogAction log_wb:=Me.Wrkbk.name _
                                , log_action:=DebugComp & "Code change backed-up (exported)"
        End If
        Set flExport = GetFile(sExportFullName)
    End If

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
        
        sTempPath = sTempFolder & "\" & vbc.name & TypeExtention(vbc)
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

Public Property Get IsHostedCommon() As Boolean
    IsHostedCommon = dctCommonHosted.Exists(sComponentName)
End Property

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
    With Me.CompOrigin
        If UpdateOriginWithUsedConfirmed Then
            '~~ Get the origin host Workbook open
            sWbHost = mCommDat.CommCompHostWorkbookFullName(sComponentName)
            .Wrkbk = mWrkbk.GetOpen(sWbHost)
            .VBComp = .Wrkbk.VBProject.VBComponents(sComponentName)
            Select Case .VBComp.Type
                Case vbext_ct_ClassModule, vbext_ct_Document, vbext_ct_StdModule
                    .SynchOriginWithUsed
                Case vbext_ct_MSForm
                    .ReplaceOriginWithUsed
            End Select
            bUpdated = True
       End If
    End With ' Me.CompOrigin

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub UpdateUsedWithOrigin()
' --------------------------------------------------
' Update is done either by code synch or by replace.
' --------------------------------------------------
    Const PROC = "UpdateUsedWithOrigin"

    On Error GoTo eh
    
    With Me ' The used class
        .ReplaceUsedWithOrigin vbc
        
        '~~ Log the updated code version
        .CodeVersionAsOfDate = .CompOrigin.ExportFile.DateLastModified
        '~~ Log the update
        mCompManLog.LogAction log_wb:=Me.Wrkbk.name _
                            , log_action:=DebugComp & "The original's last update date = " & .CompOrigin.ExportFile.DateLastModified
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
    For Each v In ExportFileToArray(Me.CompOrigin.ExportFile)
        iLine = iLine + 1
        vbcm.InsertLines iLine, v
    Next v
    Debug.Print DebugComp & Now() & " Code in " & Me.Wrkbk.name & " synchronized with code in " & Me.CompOrigin.Wrkbk.name & " ( via Export File '" & Me.CompOrigin.ExportFile.name & "'."
                
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
    Dim wb  As Workbook
                
    Debug.Print DebugComp & Now() & " About to synch code with Export File '" & Me.CompOrigin.ExportFile.name & "'."
    '~~ Remove all lines from the target component
    Set wb = mWrkbk.GetOpen(Me.CompOrigin.HostFullName)
    Set vbcm = wb.VBProject.VBComponents(sComponentName).CodeModule
    With vbcm
        .DeleteLines 1, .CountOfLines
    End With
    
    '~~ Insert/copy all code lines from the source to the target component
    iLine = 0
    For Each v In ExportFileToArray(Me.CompOrigin.ExportFile)
        iLine = iLine + 1
        vbcm.InsertLines iLine, v
    Next v
    
    Debug.Print DebugComp & Now() & " Code synchronized with Export File '" & Me.CompOrigin.ExportFile.name & "'."
                
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub ReplaceUsedWithOrigin(ByVal used_vbc As VBComponent)
' -------------------------------------------------------------
'
' -------------------------------------------------------------
    Const PROC = "ReplaceUsedWithOrigin"

    On Error GoTo eh
    
    used_vbc.name = used_vbc.name & "_Temp"
    With wb.VBProject.VBComponents
        .Import Me.CompOrigin.ExportFile.Path
        .Remove used_vbc
    End With
    Set vbc = wb.VBProject.VBComponents(sComponentName) ' renew the vbc object to point to the new imported component
    
    On Error Resume Next
    
    Application.EnableEvents = False
    wb.Save
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
    Dim wbOrigin    As Workbook
    Dim vbcOrigin   As VBComponent
    
    With Me.CompOrigin ' Me is the Active Workbook using a Common Component
        Set wbOrigin = mWrkbk.GetOpen(.HostFullName) ' Wrkbk is the Workbook hosting the Common Component
        .VBComp.name = .VBComp.name & "_temp"
        wbOrigin.VBProject.VBComponents.Import Me.ExportFile.Path
        .VBComp = wb.VBProject.VBComponents(sComponentName)
        
    End With
    Debug.Print DebugComp & Now() & " Replaced with Export File '" & Me.CompOrigin.ExportFile.name & "'."

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

    If Not wb Is Nothing _
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

Public Sub RegisterAsHostedCommon()
    mCommDat.CommCompExpFileFullName(sComponentName) = sExportFullName
    mCommDat.CommCompHostWorkbookBaseName(sComponentName) = sHostBaseName
End Sub

Public Function Exists(ByVal sComp As String) As Boolean
    Dim vbc As VBComponent

    On Error Resume Next
    Set vbc = Me.Wrkbk.VBProject.VBComponents(sComp)
    Exists = Err.Number = 0
    Set vbc = Nothing

End Function

