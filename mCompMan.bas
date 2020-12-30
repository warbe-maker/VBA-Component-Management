Attribute VB_Name = "mCompMan"
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Standard Module mCompMan: Provides all means to manage the VBComponents of
'                 another workbook but ThisWorkbook provided:
'                 - All manged Workbook resides in its own dedicated directory
'                 - The following modules are stored in is the VB-CompMan.xlsb
'                   Workbook: mCompMan, mExists, ufCompMan, clsSPP
'                 The user interface (ufCompMan) is available through the
'                 Manage method.
' Usage:          All methods are to be called from another Workbook. This
'                 module does nothing when running in ThisWorkbook.
'                 The Workbook of which the modules are to be managed is
'                 provided as Application.Run argument.
' ----------------------------------------------------------------------------
'                 Application.Run "CoMoMan.xlsb!....."
' ----------------------------------------------------------------------------
'                   ..... is the method to be executed e.g. Manage
'                   which displays a user interface.
'
' Methods:
' - ExportAll       Exports all components into the directory of
'                   the Workbook - which should be a dedicated one.
' - ImportAll       Imports all export-files into the specified
'                   Workbook. ! Needs to be executed twice !
' - InportUtdOnly   Imports only Export Files with a more recent
'                   last modified date than the Workbook.
'                   Optionally checks for a more up-to-date export
'                   files in a "Common" directory which is in the
'                   same directory as the "Workbook's directory
' - Remove          Removes a specified VBComponent from the
'                   Workbook
' - Import          Imports a single exported component. Any type
'                   document VBComponent is imported
'                   l i n e   b y   l i n e !
' - ImportByLine    Imports a single VBComponent into the specified
'                   target Workbook. The source may be an export
'                   file or another Workbook
' - Reorg           Reorganizes the code of a Workbook's component
'                   (class and standard module only!)
' - Transfer        Transfers the code of a specified VBComponent
'                   from a specified source Workbook into a
'                   specified target Workbook provided the component
'                   exists in both Workbooks
' - SynchAll        Synchronizes the code of modules existing in a
'                   target Workbook with the code in a source Workbook.
'                   Missing References are only added when the synchro-
'                   nized module is the only one in the source Workbook
'                   (i.e. the source Workbook is dedicated to this module
'                   and thus has only References required for the syn-
'                   chronized module.
' - SynchronizeFull       Synchronizes a (target) Workbook based on a template
'                   (source Workbook) in order to keep their vba code
'                   identical. Full synchronization means that modules
'                   not existing in the source Workbook are removed and
'                   modules not existing in the target Workbook are
'                   added. Missing References are added at first and
'                   obsolete Refeences are removed at last.
'                   (see SynchAll in contrast)
' - Manage          Provides all methods via a user interface.
'
' Uses Common Components: - mBasic
'                         - mErrhndlr
'                         - mFile
'                         - mVBP
'                         - mWrkbk
' Requires:
' - Reference to: - "Microsoft Visual Basic for Applications Extensibility ..."
'                 - "Microsoft Scripting Runtime"
'                 - "Windows Script Host Object Model"
'                 - Trust in the VBA project object modell (Security
'                   setting for makros)
'
' W. Rauschenberger Berlin August 2019
' -------------------------------------------------------------------------------
Public Enum enKindOfComp                ' The kind of Component in the sense of CompMan
        enUnknown = 0
        enHostedRaw = 1
        enRawClone = 2                 ' The Component is a used raw, i.e. the raw is hosted by another Workbook
        enInternal = 3             ' Neither a hosted nor a used Raw Common Component
End Enum

Public Enum enUpdateReply
    enUpdateOriginWithUsed
    enUpdateUsedWithOrigin
    enUpdateNone
End Enum

' Distinguish the code of which Workbook is allowed to be updated
Public Enum vbcmType
    vbext_ct_StdModule = 1          ' .bas
    vbext_ct_ClassModule = 2        ' .cls
    vbext_ct_MSForm = 3             ' .frm
    vbext_ct_ActiveXDesigner = 11   ' ??
    vbext_ct_Document = 100         ' .cls
End Enum

Public Enum enKindOfCodeChange  ' Kind of code change
    enUnknown
    enUsedOnly              ' A component which is neither a hosted raw nor a raw's clone has changed
    enRawOnly               ' Only the remote raw code hosted in another Workbook had changed
    enCloneOnly             ' Only the code of the target Component had changed
    enRawAndClone           ' Both, the code of the remote raw and the raw clone had changed
    enNoCodeChange          ' No code change at all
    enPendingExportOnly     ' A modified raw may have been re-imported already
End Enum

Public cComp            As clsComp
Public cRaw             As clsRaw
Public cLog             As clsLog
Public asNoSynch()      As String
Public dctRawComponents As Dictionary
Public dctRawHosts      As Dictionary

Private dctHostedRaws   As Dictionary
Private lMaxCompLength  As Long

Public Property Get HostedRaws() As Variant
    Set HostedRaws = dctHostedRaws
End Property

Private Property Let HostedRaws(ByVal vhosted As Variant)
' -------------------------------------------------------
' Saves the names of the hosted raw components
' (hosted_raws) to the Dictionary (dctRawComponents).
' -------------------------------------------------
    Dim v               As Variant
    Dim sComp           As String
    
    If dctHostedRaws Is Nothing Then Set dctHostedRaws = New Dictionary
    For Each v In Split(vhosted, ",")
        sComp = Trim$(v)
        If Not dctHostedRaws.Exists(sComp) Then
            dctHostedRaws.Add sComp, sComp
        End If
    Next v
    
End Property

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mCompMan" & "." & es_proc
End Function

Public Sub ExportAll(Optional ByVal exp_wrkbk As Workbook = Nothing)
' -----------------------------------------------------------
'
' -----------------------------------------------------------
    Const PROC = "ExportAll"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    
    mErH.BoP ErrSrc(PROC)
    
    If exp_wrkbk Is Nothing Then Set exp_wrkbk = ActiveWorkbook
    Set cComp = New clsComp
    
    With cComp
        If wbAddIn.IsAddinInstance _
        Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The Workbook (active or provided) is the CompMan Addin instance which is impossible for this operation!"
        .Wrkbk = exp_wrkbk
        For Each vbc In .Wrkbk.VBProject.VBComponents
            .VBComp = vbc
            .BackUpCode
        Next vbc
    End With

xt: Set cComp = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub ExportChangedComponents(ByVal ec_wb As Workbook, _
                          Optional ByVal ec_hosted As String = vbNullString)
' ------------------------------------------------------------------------
' Exclusively performed/trigered by the Before_Save event:
' - Any code change (detected by the comparison of a temporary export file
'   with the current export file) is backed-up, i.e. exported
' - Any Export Files representing no longer existing components are removed
' - In case of conflicting or unusual code modifications in a raw clone
'   component the user decides what to do with it. Choices may be:
'   -- Modifications are ignored, i.e. will be reverted with the next open
'   -- The raw is updated, i.e. the modifications in the clone become
'      common for all VB-Projects using a clone of this raw
'   -- The user merges those modifications desired for becoming common
'      and ignores others.
' Background:
' - This procedure is preferrably triggered by the Before_Save event.
' - The ExportFile's last access date reflects the date of the last code
'   change. This date is logged when a used Common Component is updated.
' ------------------------------------------------------------------------
    Const PROC = "ExportChangedComponents"
    
    On Error GoTo eh
    Dim lCompMaxLen         As Long
    Dim sComp               As String
    Dim vbc                 As VBComponent
    Dim dctComps            As Dictionary
    Dim dctRemove           As Dictionary
    Dim v                   As Variant
    Dim fl                  As FILE
    Dim sFolder             As String
    Dim lComponents         As Long
    Dim lExported           As Long
    Dim lRemoved            As Long
    Dim sExported           As String
    Dim bUpdated            As Boolean
    Dim lUpdated            As Long
    Dim sUpdated            As String
    Dim sMsg                As String
    Dim fso                 As New FileSystemObject
    Dim sServiced           As String
    Dim sProgress           As String
    
    mErH.BoP ErrSrc(PROC)
    
    If ec_wb Is Nothing Then Set ec_wb = ActiveWorkbook
    '~~ Prevent any action for a Workbook opened with any irregularity
    If InStr(ActiveWindow.caption, "(") <> 0 Then GoTo xt
    If InStr(ec_wb.FullName, "(") <> 0 Then GoTo xt
    
    lCompMaxLen = MaxCompLength(wb:=ec_wb)
    Set dctComps = New Dictionary
    Set cComp = New clsComp
    Set cLog = New clsLog
    cLog.Reset
    cLog.Service = ErrSrc(PROC)
    
    MaintainHostedRaws mh_hosted:=ec_hosted _
                     , mh_wb:=ec_wb
    
    sProgress = String$(ec_wb.VBProject.VBComponents.Count + 1, ".")
    For Each vbc In ec_wb.VBProject.VBComponents
        sProgress = Left(sProgress, Len(sProgress) - 1)
        Application.StatusBar = "Export of changed components: " & Format(lExported, "##") & sProgress
        mTrc.BoC ErrSrc(PROC) & " " & vbc.name
        Set cComp = New clsComp
        With cComp
            .Wrkbk = ec_wb
            .VBComp = vbc
            .ComponentName = vbc.name
            sServiced = .Wrkbk.name & " Component """ & vbc.name & """"
            sServiced = sServiced & String(lCompMaxLen - Len(vbc.name), ".")
            cLog.Serviced = sServiced
            If .CodeModuleIsEmpty Then GoTo next_vbc
        End With
        
        lComponents = lComponents + 1
        dctComps.Add vbc.name, vbc
        
        Select Case cComp.KindOfComp
            Case enInternal, enHostedRaw
                '~~ This is a raw component's clone
                Select Case cComp.KindOfCodeChange
                    Case enCloneOnly, enPendingExportOnly, enRawAndClone, enRawOnly, enUsedOnly
                        mTrc.BoC ErrSrc(PROC) & " Backup No-Raw " & vbc.name
                        cComp.BackUpCode
                        lExported = lExported + 1
                        sExported = vbc.name & ", " & sExported
                        mTrc.EoC ErrSrc(PROC) & " Backup No-Raw" & vbc.name
                        GoTo next_vbc
                End Select
                
                If cComp.KindOfComp = enHostedRaw Then
                    If mRaw.ExpFileFullName(comp_name:=cComp.CompName) <> cComp.ExportFileFullName Then
                        mRaw.ExpFileFullName(comp_name:=cComp.CompName) = cComp.ExportFileFullName
                        cLog.Action = "Component's Export File Full Name registered"
                    End If
                End If

            Case enRawClone
                '~~ Establish a component class object which represents the cloned raw's remote instance
                '~~ which is hosted in another Workbook
                Set cRaw = New clsRaw
                With cRaw ' Provide all available information rearding the remote raw component
                    .CompName = cComp.CompName
                    .ExpFile = fso.GetFile(mRaw.ExpFileFullName(comp_name:=.CompName))
                    .ExpFileFullName = .ExpFile.PATH
                    .HostFullName = mRaw.HostFullName(comp_name:=.CompName)
                End With
                
                With cComp
                    Select Case .KindOfCodeChange
                        Case enPendingExportOnly
                            mTrc.BoC ErrSrc(PROC) & " Backup Clone " & vbc.name
                            cComp.BackUpCode
                            lExported = lExported + 1
                            sExported = vbc.name & ", " & sExported
                            mTrc.EoC ErrSrc(PROC) & " Backup Clone" & vbc.name
                            GoTo next_vbc
                        
                        Case enNoCodeChange
                            cLog.Action = "No action performed"

                        Case enRawAndClone
                            '~~ The user will decide which of the code modification will go to the raw and the raw will be
                            '~~ updated with the final result
                            cLog.Action = "No action performed"
                            
                        Case enCloneOnly
                            '~~ This is regarded an unusual code change because instead of maintaining the origin code
                            '~~ of a Common Component in ist "host" VBProject it had been changed in the using VBProject.
                            '~~ Nevertheless updating the origin code with this change is possible when explicitely confirmed.
                            mFile.Compare file_left_full_name:=cComp.ExportFileFullName _
                                      , file_left_title:=cComp.ExportFileFullName _
                                      , file_right_full_name:=cRaw.ExpFileFullName _
                                      , file_right_title:=cRaw.ExpFileFullName
                                      
                            .ReplaceRemoteWithClonedRawWhenConfirmed rwu_updated:=bUpdated ' when confirmed in user dialog
                            If bUpdated Then
                                lUpdated = lUpdated + 1
                                sUpdated = vbc.name & ", " & sUpdated
                                cLog.Action = """Remote Raw"" has been updated with code of ""Raw Clone"""
                            End If
                        Case enRawOnly
                            Debug.Print "Remote raw " & vbc.name & " has changed and will update the used in the VB-Project the next time it is opened."
                            '~~ The changed remote raw will be used to update the clone the next time the Workbook is openend
                            cLog.Action = "No action performed"
                        Case enUsedOnly
                            cLog.Action = "No action performed"
                    End Select
                End With
        End Select
                                
next_vbc:
        Set cComp = Nothing
        Set cRaw = Nothing
        mTrc.EoC ErrSrc(PROC) & " " & vbc.name
    Next vbc
        
    '~~ Remove Export Files of no longer existent components
    Set dctRemove = New Dictionary
    Set cComp = New clsComp
    With cComp
        .Wrkbk = ec_wb
        sFolder = .ExportFolder
    End With
    
    With fso
        For Each fl In .GetFolder(sFolder).Files
            Select Case .GetExtensionName(fl.PATH)
                Case "bas", "cls", "frm", "frx"
                    sComp = .GetBaseName(fl.PATH)
                    If Not dctComps.Exists(sComp) Then
                        dctRemove.Add fl, fl.name
                    End If
            End Select
        Next fl
    
        For Each v In dctRemove
            .DeleteFile v.PATH
            lRemoved = lRemoved + 1
        Next v
    End With
    Set cComp = Nothing

    Select Case lExported
        Case 0:     sMsg = "None of the " & lComponents & " Components in Workbook " & ec_wb.name & " had been changed/exported/backed up."
        Case 1:     sMsg = " 1 Component (of " & lComponents & ") in Workbook " & ec_wb.name & " had been exported/backed up: " & Left(sExported, Len(sExported) - 2)
        Case Else:  sMsg = lExported & " Components (of " & lComponents & ") in Workbook " & ec_wb.name & " had been exported/backed up: " & Left(sExported, Len(sExported) - 1)
    End Select
    If Len(sMsg) > 255 Then sMsg = Left(sMsg, 251) & " ..."
    Application.StatusBar = sMsg
    
xt: mErH.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub MaintainHostedRaws(ByVal mh_hosted As String, _
                               ByVal mh_wb As Workbook)
' ---------------------------------------------------------
'
' ---------------------------------------------------------
    Const PROC = "MaintainHostedRaws"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim fso As New FileSystemObject
    
    mErH.BoP ErrSrc(PROC)

    HostedRaws = mh_hosted
    If HostedRaws.Count <> 0 Then
        If Not mHost.Exists(raw_host_base_name:=fso.GetBaseName(mh_wb.FullName)) _
        Or mHost.FullName(host_base_name:=fso.GetBaseName(mh_wb.FullName)) <> mh_wb.FullName Then
            '~~ Keep a record when this Workbook hosts one or more Raw components and not is already registered
            mHost.FullName(host_base_name:=fso.GetBaseName(mh_wb.FullName)) = mh_wb.FullName
            cLog.Action = "Workbook registered as a host for at least one raw component"
        End If
    
        For Each v In HostedRaws
            '~~ Keep a record for each of the raw components hosted by this Workbook
            If Not mRaw.Exists(raw_comp_name:=v) _
            Or mRaw.HostFullName(comp_name:=v) <> mh_wb.FullName Then
                mRaw.HostFullName(comp_name:=v) = mh_wb.FullName
                cLog.Action = "Raws components's host Workbook registered"
            End If
        Next v
    Else
        '~~ Remove any raws still existing and pointing to this Workbook as host
        For Each v In mRaw.Components
            If mRaw.HostFullName(comp_name:=v) = mh_wb.FullName Then
                mRaw.Remove comp_name:=v
                cLog.Action = "Component removed from '" & mHost.DAT_FILE & "'"
            End If
        Next v
        If mHost.Exists(fso.GetBaseName(mh_wb.FullName)) Then
            mHost.Remove (fso.GetBaseName(mh_wb.FullName))
            cLog.Action = "Workbook no longer a host for at least one raw component removed from '" & mHost.DAT_FILE & "'"
        End If
    End If

xt: Set fso = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub UpdateClonesTheRawHasChanged( _
                                  ByVal uc_wb As Workbook, _
                         Optional ByVal uc_hosted As String = vbNullString)
' -------------------------------------------------------------------------------
' Updates any cloned raw component in the target Workbook (updt_wrkbk) where the
' code of the remote raw component hosted in another Workbook has changed (the
' export files differ).
' -------------------------------------------------------------------------------
    Const PROC = "UpdateClonesTheRawHasChanged"
    
    On Error GoTo eh
    Dim wb                  As Workbook
    Dim vbc                 As VBComponent
    Dim lCompMaxLen         As Long
    Dim lComponents         As Long
    Dim lClonedRaw          As Long
    Dim lReplaced           As Long
    Dim sReplaced           As String
    Dim sMsg                As String
    Dim sServiced           As String
    Dim fso                 As New FileSystemObject
    
    mErH.BoP ErrSrc(PROC)
    '~~ Prevent any action for a Workbook opened with any irregularity
    If InStr(ActiveWindow.caption, "(") <> 0 Then GoTo xt
    If InStr(uc_wb.FullName, "(") <> 0 Then GoTo xt
    Set wb = uc_wb
    
    MaintainHostedRaws mh_hosted:=uc_hosted _
                     , mh_wb:=uc_wb
    
    lCompMaxLen = MaxCompLength(wb:=wb)
    Set cLog = New clsLog
    cLog.Service = ErrSrc(PROC)
    
    For Each vbc In wb.VBProject.VBComponents
        Set cComp = New clsComp
        cComp.Wrkbk = wb
        cComp.VBComp = vbc
        lComponents = lComponents + 1
        sServiced = cComp.Wrkbk.name & " Component """ & vbc.name & """"
        sServiced = sServiced & String(lCompMaxLen - Len(vbc.name), ".")
        cLog.Serviced = sServiced
            
        If cComp.KindOfComp = enRawClone Then
            '~~ Establish a component class object which represents the cloned raw's remote instance
            '~~ which is hosted in another Workbook
            Set cRaw = New clsRaw
            With cRaw ' Provide all available information rearding the remote raw component
                .CompName = cComp.CompName
                .ExpFile = fso.GetFile(FilePath:=mRaw.ExpFileFullName(.CompName))
                .ExpFileFullName = .ExpFile.PATH
                .HostFullName = mRaw.HostFullName(comp_name:=.CompName)
            End With

            With cComp
                If .KindOfComp = enRawClone Then lClonedRaw = lClonedRaw + 1
                If .KindOfCodeChange = enRawOnly _
                Or .KindOfCodeChange = enRawAndClone Then
                    '~~ Attention!! The cloned raw's code is updated disregarding any code changed in it.
                    '~~ A code change in the cloned raw is only considered when the Workbook is about to
                    '~~ be closed - where it may be ignored to make exactly this happens.
                    .ReplaceClonedWithRemoteRaw ru_vbc:=vbc
                    lReplaced = lReplaced + 1
                    sReplaced = .CompName & ", " & sReplaced
                    '~~ Register the update being used to identify a potentially relevant
                    '~~ change of the origin code
                    cLog.Action = "Replaced! The remote raw's export file '" & cRaw.ExpFileFullName & "' has been imported."
                    .BackUpCode
                End If
            End With
        End If
        Set cComp = Nothing
    Next vbc
        
    Select Case lClonedRaw
        Case 0: sMsg = "None of " & lComponents & " had been identified as ""Cloned Raw Component""."
        Case 1
            Select Case lReplaced
                Case 0:     sMsg = "1 of " & lComponents & " has been identified as ""Cloned Raw Component"" but has not been updated since the raw had not changed."
                Case 1:     sMsg = "1 Component of " & lComponents & " has been identified as ""Cloned Raw Component"" and updated because the raw had changed (" & Left(sReplaced, Len(sReplaced) - 2) & ")."
               End Select
        Case Else
            Select Case lReplaced
                Case 0:     sMsg = lClonedRaw & " Components of " & lComponents & " had been identified as ""Cloned Raw Components"". None had been updated since none of the raws had changed."
                Case 1:     sMsg = lClonedRaw & " Components of " & lComponents & " had been identified as ""Cloned Raw Components"". One had been updated since the raw's code had changed (" & Left(sReplaced, Len(sReplaced) - 2) & ")."
                Case Else:  sMsg = lClonedRaw & " Components of " & lComponents & " had been identified as ""Cloned Raw Components"". " & lReplaced & " have been updated because the raw's code had changed (" & Left(sReplaced, Len(sReplaced) - 2) & ")."
            End Select
    End Select
    If Len(sMsg) > 255 Then sMsg = Left(sMsg, 251) & " ..."
    Application.StatusBar = sMsg
    
xt: Set cRaw = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Version(ByVal c_version As clsAddinVersion)
' ---------------------------------------------------------------------------------------------------------------------
' Called by the development instance via Application.Run. Because the version value cannot be returned to the call via
' a ByRef argument, a class object is used instead.
' See: http://www.tushar-mehta.com/publish_train/xl_vba_cases/1022_ByRef_Argument_with_the_Application_Run_method.shtml
' ---------------------------------------------------------------------------------------------------------------------
    c_version.Version = wbAddIn.AddInVersion
End Sub

Public Property Get MaxCompLength(Optional ByVal wb As Workbook) As Long
    Dim vbc As VBComponent
    If lMaxCompLength = 0 Then
        For Each vbc In wb.VBProject.VBComponents
            lMaxCompLength = mBasic.Max(lMaxCompLength, Len(vbc.name))
        Next vbc
    End If
    MaxCompLength = lMaxCompLength
End Property

Public Sub CompareCloneWithRaw(ByVal cmp_comp_name As String)
' -----------------------------------------------------------
'
' -----------------------------------------------------------
    Const PROC = "CompareCloneWithRaw"
    
    On Error GoTo eh
    Dim sExpFileClone   As String
    Dim sExpFileRaw     As String
    Dim wb              As Workbook
    Dim cComp           As New clsComp
    
    Set wb = ActiveWorkbook
    With cComp
        .Wrkbk = wb
        .CompName = cmp_comp_name
        .VBComp = wb.VBProject.VBComponents(.CompName)
        sExpFileRaw = mRaw.ExpFileFullName(comp_name:=cmp_comp_name)
        sExpFileClone = .ExportFileFullName
    
        mFile.Compare file_left_full_name:=sExpFileClone _
                    , file_right_full_name:=sExpFileRaw _
                    , file_left_title:="The cloned raw's current code in Workbook/VBProject " & cComp.WrkbkBaseName & " (" & sExpFileClone & ")" _
                    , file_right_title:="The remote raw's current code in Workbook/VBProject " & mBasic.BaseName(mRaw.HostFullName(.CompName)) & " (" & sExpFileRaw & ")"

    End With
    Set cComp = Nothing

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub DisplayCodeChange(ByVal cmp_comp_name As String)
' -----------------------------------------------------------
'
' -----------------------------------------------------------
    Const PROC = "DisplayCodeChange"
    
    On Error GoTo eh
    Dim sExpFileTemp    As String
    Dim wb              As Workbook
    Dim cComp           As New clsComp
    Dim fso             As New FileSystemObject
    Dim sTempFolder     As String
    Dim flExpTemp       As FILE
    
    Set wb = ActiveWorkbook
    With cComp
        .Wrkbk = wb
        .CompName = cmp_comp_name
        .VBComp = wb.VBProject.VBComponents(.CompName)
    End With
    
    With fso
        sTempFolder = .GetFile(cComp.ExportFileFullName).ParentFolder & "\Temp"
        If Not .FolderExists(sTempFolder) Then .CreateFolder sTempFolder
        sExpFileTemp = sTempFolder & "\" & cComp.CompName & cComp.TypeExtention(cComp.VBComp)
        cComp.VBComp.Export sExpFileTemp
        Set flExpTemp = .GetFile(sExpFileTemp)
    End With

    With cComp
        mFile.Compare file_left_full_name:=sExpFileTemp _
                    , file_right_full_name:=.ExportFileFullName _
                    , file_left_title:="The component's current code in Workbook/VBProject " & cComp.WrkbkBaseName & " ('" & sExpFileTemp & "')" _
                    , file_right_title:="The component's currently exported code in '" & .ExportFileFullName & "'"

    End With
    
xt: If fso.FolderExists(sTempFolder) Then fso.DeleteFolder (sTempFolder)
    Set cComp = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Merge(Optional ByVal fl_1 As String = vbNullString, _
                 Optional ByVal fl_2 As String = vbNullString)
' -----------------------------------------------------------
'
' -----------------------------------------------------------
    Const PROC = "Merge"
    
    On Error GoTo eh
    Dim fl_left         As FILE
    Dim fl_right        As FILE
    
    If fl_1 = vbNullString Then mFile.SelectFile sel_result:=fl_left
    If fl_2 = vbNullString Then mFile.SelectFile sel_result:=fl_right
    fl_1 = fl_left.PATH
    fl_2 = fl_right.PATH
    
    mFile.Compare file_left_full_name:=fl_1 _
                , file_right_full_name:=fl_2 _
                , file_left_title:=fl_1 _
                , file_right_title:=fl_2

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

