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
        enKoCunknown = 0
        enRawRemote = 1                 ' The Component is a hosted raw one
        enRawHosted = 2
        enRawClone = 3                 ' The Component is a used raw, i.e. the raw is hosted by another Workbook
        enNoRaw = 4                     ' Neither a hosted nor a used Raw Common Component
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
    enKindUnknown
    enUsedOnly                  ' A component which is neither a remote raw nor a cloned raw has changed
    enRawRemoteOnly             ' Only the remote raw code hosted in another Workbook had changed
    enRawCloneOnly              ' Only the code of the target Component had changed
    enRawCloneAndRemote         ' Both, the code of the remote raw and the raw clone had changed
    enNoCodeChange              ' No code change at all
End Enum

Public dat          As clsAppData
Public wblog        As clsAppData
Private wbTarget    As Workbook     ' The Workbook of which the VB-Components are managed
                                    ' (the Workbook which "pulls" its up to date VB-Project from wbSource)
Public cComp        As clsComp
Public cRaw         As clsRaw
Public cLog         As clsLog
Public asNoSynch()  As String
Public dctRaws      As Dictionary
Public dctRawHosts  As Dictionary

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

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub ExportChangedComponents(ByVal cc_wb As Workbook, _
                          Optional ByVal cc_hosted As String = vbNullString)
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
    Dim flRawExportFile     As FILE
    Dim lComponents         As Long
    Dim lExported           As Long
    Dim lRemoved            As Long
    Dim sExported           As String
    Dim bUpdated            As Boolean
    Dim lUpdated            As Long
    Dim sUpdated            As String
    Dim sMsg                As String
    Dim lKindOfComp         As enKindOfComp
    Dim flRawExpFile        As FILE
    Dim sRawHostFullName    As String
    Dim fso                 As New FileSystemObject
    Dim sServiced           As String
    
    mErH.BoP ErrSrc(PROC)
    
    If cc_wb Is Nothing Then Set cc_wb = ActiveWorkbook
    '~~ Prevent any action for a Workbook opened with any irregularity
    If InStr(ActiveWindow.Caption, "(") <> 0 Then GoTo xt
    If InStr(cc_wb.FullName, "(") <> 0 Then GoTo xt
    
    Set dctComps = New Dictionary
    Set cComp = New clsComp
    Set cLog = New clsLog
    cLog.Reset
    cLog.Service = ErrSrc(PROC)
    
    lCompMaxLen = mDat.CommCompsMaxLenght()
    If cc_hosted <> vbNullString Then
        '~~ Keep a record when this Workbook hosts one or more Common Components
        mDat.CommCompsHostWorkbookFullName(fso.GetBaseName(cc_wb.FullName)) = cc_wb.FullName
    End If
    
    For Each vbc In cc_wb.VBProject.VBComponents
        Application.StatusBar = "Export of changed components: " & Format(lExported, "##") & String$((cc_wb.VBProject.VBComponents.Count + 1) - lComponents - lExported, ".")
        mTrc.BoC ErrSrc(PROC) & " " & vbc.name
        Set cComp = New clsComp
        With cComp
            .Wrkbk = cc_wb
            .HostedRawComps = cc_hosted
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
            Case enNoRaw ' This is a 'normal' component which is neither a hosted nor a cloned raw component
                If Not cComp.KindOfCodeChange = enNoCodeChange Then
                    mTrc.BoC ErrSrc(PROC) & " Backup No-Raw " & vbc.name
                    cComp.BackUpCode
                    lExported = lExported + 1
                    sExported = vbc.name & ", " & sExported
                    mTrc.EoC ErrSrc(PROC) & " Backup No-Raw" & vbc.name
                    cLog.Action = "Exported to '" & cComp.ExportFileFullName & "'"
                    GoTo next_vbc
                End If

            Case enRawClone
                '~~ Establish a component class object which represents the cloned raw's remote instance
                '~~ which is hosted in another Workbook
                Set cRaw = New clsRaw
                mDat.IsRaw raw_comp_name:=cComp.CompName, raw_exp_file:=flRawExpFile, raw_host_full_name:=sRawHostFullName
                With cRaw ' Provide all available information rearding the remote raw component
                    .ExpFile = flRawExpFile
                    .HostFullName = sRawHostFullName
                    .CompName = cComp.CompName
                End With
                
                With cComp
                    Select Case .KindOfCodeChange
                        Case enNoCodeChange
                            cLog.Action = "No action performed"

                        Case enRawCloneAndRemote
                            '~~ The user will decide which of the code modification will go to the raw and the raw will be
                            '~~ updated with the final result
                            cLog.Action = "No action performed"
                            
                        Case enRawCloneOnly
                            '~~ This is regarded an unusual code change because instead of maintaining the origin code
                            '~~ of a Common Component in ist "host" VBProject it had been changed in the using VBProject.
                            '~~ Nevertheless updating the origin code with this change is possible when explicitely confirmed.
                            mFl.Compare file_left_full_name:=cComp.ExportFileFullName _
                                      , file_left_title:=cComp.ExportFileFullName _
                                      , file_right_full_name:=cRaw.ExpFileFullName _
                                      , file_right_title:=cRaw.ExpFileFullName
                                      
                            .ReplaceRemoteWithClonedRawWhenConfirmed rwu_updated:=bUpdated ' when confirmed in user dialog
                            If bUpdated Then
                                lUpdated = lUpdated + 1
                                sUpdated = vbc.name & ", " & sUpdated
                                cLog.Action = """Remote Raw"" has been updated with code of ""Raw Clone"""
                            End If
                        Case enRawRemoteOnly
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
    cComp.Wrkbk = cc_wb
    sFolder = cComp.ExportFolder
    With New FileSystemObject
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
        Case 0:     sMsg = "None of the " & lComponents & " Components in Workbook " & cc_wb.name & " had been changed/exported/backed up."
        Case 1:     sMsg = " 1 Component (of " & lComponents & ") in Workbook " & cc_wb.name & " had been exported/backed up: " & left(sExported, Len(sExported) - 2)
        Case Else:  sMsg = lExported & " Components (of " & lComponents & ") in Workbook " & cc_wb.name & " had been exported/backed up: " & left(sExported, Len(sExported) - 1)
    End Select
    If Len(sMsg) > 255 Then sMsg = left(sMsg, 251) & " ..."
    Application.StatusBar = sMsg
    
xt: mErH.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub UpdateClonesTheRawHasChanged( _
                                  ByVal cc_wb As Workbook, _
                         Optional ByVal updt_hosted_raws As String = vbNullString)
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
    Dim flRawExportFile     As FILE
    Dim sRawHostFullName    As String
    Dim lComponents         As Long
    Dim lClonedRaw          As Long
    Dim lReplaced           As Long
    Dim sReplaced           As String
    Dim sMsg                As String
    Dim lKindOfComp         As enKindOfComp
    Dim sServiced           As String

    mErH.BoP ErrSrc(PROC)
    '~~ Prevent any action for a Workbook opened with any irregularity
    If InStr(ActiveWindow.Caption, "(") <> 0 Then GoTo xt
    If InStr(ActiveWorkbook.FullName, "(") <> 0 Then GoTo xt
    Set wb = ActiveWorkbook
    
    lCompMaxLen = mDat.CommCompsMaxLenght()
    Set cLog = New clsLog
    cLog.Service = ErrSrc(PROC)
    
    For Each vbc In wb.VBProject.VBComponents
        Set cComp = New clsComp
        cComp.Wrkbk = wb
        cComp.HostedRawComps = updt_hosted_raws
        cComp.VBComp = vbc
        lComponents = lComponents + 1
        sServiced = cComp.Wrkbk.name & " Component """ & vbc.name & """"
        sServiced = sServiced & String(lCompMaxLen - Len(vbc.name), ".")
        cLog.Serviced = sServiced
            
        If cComp.KindOfComp = enRawClone Then
            '~~ Establish a component class object which represents the cloned raw's remote instance
            '~~ which is hosted in another Workbook
            Set cRaw = New clsRaw
            mDat.IsRaw raw_comp_name:=cComp.CompName, raw_exp_file:=flRawExportFile, raw_host_full_name:=sRawHostFullName
            With cRaw ' Provide all available information rearding the remote raw component
                .ExpFile = flRawExportFile
                .HostFullName = sRawHostFullName
                .CompName = cComp.CompName
            End With

            With cComp
                If .KindOfComp = enRawClone Then lClonedRaw = lClonedRaw + 1
                If .KindOfCodeChange = enRawRemoteOnly _
                Or .KindOfCodeChange = enRawCloneAndRemote Then
                    '~~ Attention!! The cloned raw's code is updated disregarding any code changed in it.
                    '~~ A code change in the cloned raw is only considered when the Workbook is about to
                    '~~ be closed - where it may be ignored to make exactly this happens.
                    .ReplaceClonedWithRemoteRaw ru_vbc:=vbc
                    lReplaced = lReplaced + 1
                    sReplaced = .CompName & ", " & sReplaced
                    '~~ Register the update being used to identify a potentially relevant
                    '~~ change of the origin code
                    .CodeVersionAsOfDate = .ExpFile.DateLastModified
                    cLog.Action = "Replaced (re-imported) with the remote raw's export file '" & cRaw.ExpFileFullName & "'"
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
                Case 1:     sMsg = "1 Component of " & lComponents & " has been identified as ""Cloned Raw Component"" and updated because the raw had changed (" & left(sReplaced, Len(sReplaced) - 2) & ")."
               End Select
        Case Else
            Select Case lReplaced
                Case 0:     sMsg = lClonedRaw & " Components of " & lComponents & " had been identified as ""Cloned Raw Components"". None had been updated since none of the raws had changed."
                Case 1:     sMsg = lClonedRaw & " Components of " & lComponents & " had been identified as ""Cloned Raw Components"". One had been updated since the raw's code had changed (" & left(sReplaced, Len(sReplaced) - 2) & ")."
                Case Else:  sMsg = lClonedRaw & " Components of " & lComponents & " had been identified as ""Cloned Raw Components"". " & lReplaced & " have been updated because the raw's code had changed (" & left(sReplaced, Len(sReplaced) - 2) & ")."
            End Select
    End Select
    If Len(sMsg) > 255 Then sMsg = left(sMsg, 251) & " ..."
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

