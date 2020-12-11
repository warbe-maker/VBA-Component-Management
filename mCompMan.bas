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
        enRemoteRaw = 1                 ' The Component is a hosted raw one
        enHostedRaw = 2
        enClonedRaw = 3                 ' The Component is a used raw, i.e. the raw is hosted by another Workbook
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
Public cRemoteRaw   As clsComp
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
' - Exports/backs up any Component the code differs from its corresponding
'   backup/export file.
' - Exports/backs up any Component which never had been exported before
' - Removes all Export Files representing no longer existing Components.
' - Requests confirmation to update the origin code for any used Common
'   Components of which the code has been changed within the using
'   VBProject.
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
    Dim lKindOfChange       As enKindOfCodeChange
    Dim flRawExpFile        As FILE
    Dim sRawHostFullName    As String
    
    mErH.BoP ErrSrc(PROC)
    
    If cc_wb Is Nothing Then Set cc_wb = ActiveWorkbook
    '~~ Prevent any action for a Workbook opened with any irregularity
    If InStr(ActiveWindow.caption, "(") <> 0 Then GoTo xt
    If InStr(cc_wb.FullName, "(") <> 0 Then GoTo xt
    
    Set dctComps = New Dictionary
    Set cComp = New clsComp
    
    cComp.Wrkbk = cc_wb
    cComp.HostedRawComps = cc_hosted
    lCompMaxLen = mDat.CommCompsMaxLenght()
    If cc_hosted <> vbNullString Then cComp.RegisterAsHostWorkbook
        
        '~~ Keep a record when this Workbook hosts one or more Common Components
        
    For Each vbc In cc_wb.VBProject.VBComponents
        With cComp
            .VBComp = vbc
            .ComponentName = vbc.name
            If .CodeModuleIsEmpty Then GoTo next_vbc
        End With
        
        lComponents = lComponents + 1
        dctComps.Add vbc.name, vbc
        
        Select Case cComp.KindOfComp
            Case enNoRaw ' This is a 'normal' component which is neither a hosted nor a cloned raw component
                If cComp.CodeChanged Then
                    cComp.BackUpCode
                    lExported = lExported + 1
                    sExported = vbc.name & ", " & sExported
                    GoTo next_vbc
                End If

            Case enClonedRaw
                '~~ Establish a component class object which represents the cloned raw's remote instance
                '~~ which is hosted in another Workbook
                Set cRemoteRaw = New clsComp
                mDat.IsRaw raw_comp_name:=cComp.CompName, raw_exp_file:=flRawExpFile, raw_host_full_name:=sRawHostFullName
                With cRemoteRaw ' Provide all available information rearding the remote raw component
                    .ExportFile = flRawExpFile
                    .WrkbkFullName = sRawHostFullName
                    .CompName = cComp.CompName
                    .CodeVersionAsOfDate = .ExportFile.DateLastModified
                End With
                
                If cComp.KindOfCodeChange = enRawCloneOnly Then
                    '~~ This is regarded an unusual code change because instead of maintaining the origin code
                    '~~ of a Common Component in ist "host" VBProject it had been changed in the using VBProject.
                    '~~ Nevertheless updating the origin code with this change is possible when explicitely confirmed.
                    cComp.ReplaceRemoteWithClonedRawWhenConfirmed rwu_updated:=bUpdated ' when confirmed in user dialog
                    If bUpdated Then
                        lUpdated = lUpdated + 1
                        sUpdated = vbc.name & ", " & sUpdated
                    End If
                End If
        End Select
                
        
next_vbc:
    Next vbc
        
        '~~ Remove Export Files of no longer existent components
        Set dctRemove = New Dictionary
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
    
        Select Case lExported
            Case 0:     sMsg = "None of the " & lComponents & " Components in Workbook " & cc_wb.name & " had been changed/exported/backed up."
            Case 1:     sMsg = " 1 Component (of " & lComponents & ") in Workbook " & cc_wb.name & " had been exported/backed up: " & Left(sExported, Len(sExported) - 2)
            Case Else:  sMsg = lExported & " Components (of " & lComponents & ") in Workbook " & cc_wb.name & " had been exported/backed up: " & Left(sExported, Len(sExported) - 1)
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

Public Sub UpdateClonedRawsTheRemoteRawHasChanged( _
           ByVal updt_wrkbk As Workbook, _
  Optional ByVal updt_hosted_raws As String = vbNullString)
' ---------------------------------------------------------
' Updates any cloned raw component in the target Workbook
' (updt_wrkbk) where the code of the remote raw component
' hosted in another Workbook has changed (the export
' files differ).
' ---------------------------------------------------------
    Const PROC = "UpdateClonedRawsTheRemoteRawHasChanged"
    
    On Error GoTo eh
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
    Dim lKindOfChange       As enKindOfCodeChange

    mErH.BoP ErrSrc(PROC)
    '~~ Prevent any action for a Workbook opened with any irregularity
    If InStr(ActiveWindow.caption, "(") <> 0 Then GoTo xt
    If InStr(updt_wrkbk.FullName, "(") <> 0 Then GoTo xt
    
    Set cComp = New clsComp
    
    cComp.Wrkbk = updt_wrkbk
    cComp.HostedRawComps = updt_hosted_raws
    lCompMaxLen = mDat.CommCompsMaxLenght()
    
    For Each vbc In updt_wrkbk.VBProject.VBComponents
        cComp.VBComp = vbc
        lComponents = lComponents + 1
            
        If cComp.KindOfComp = enClonedRaw Then
            '~~ Establish a component class object which represents the cloned raw's remote instance
            '~~ which is hosted in another Workbook
            Set cRemoteRaw = New clsComp
            mDat.IsRaw raw_comp_name:=cComp.CompName, raw_exp_file:=flRawExportFile, raw_host_full_name:=sRawHostFullName
            With cRemoteRaw ' Provide all available information rearding the remote raw component
                .ExportFile = flRawExportFile
                .WrkbkFullName = sRawHostFullName
                .CompName = cComp.CompName
                .CodeVersionAsOfDate = .ExportFile.DateLastModified
            End With

            If cComp.KindOfCodeChange = enRawRemoteOnly _
            Or cComp.KindOfCodeChange = enRawCloneAndRemote Then
                '~~ Attention!! The cloned raw's code is updated disregarding any code changed in it.
                '~~ A code change in the cloned raw is only considered when the Workbook is about to
                '~~ be closed - where it may be ignored to make exactly this happens.
                lClonedRaw = lClonedRaw + 1
                cComp.ReplaceClonedWithRemoteRaw ru_vbc:=vbc
                lReplaced = lReplaced + 1
                sReplaced = cComp.CompName & ", " & sReplaced
                '~~ Register the update being used to identify a potentially relevant
                '~~ change of the origin code
                cComp.CodeVersionAsOfDate = cComp.ExportFile.DateLastModified
                cComp.BackUpCode
            End If
        End If
    Next vbc
        
    Select Case lClonedRaw
        Case 0: sMsg = "No Components of " & lComponents & " had been identified as ""Common Used ""."
        Case 1
            Select Case lReplaced
                Case 0:     sMsg = "1 Component of " & lComponents & " has been identified as ""Used Common Component "" but not updated since the code origin had not changed."
                Case 1:     sMsg = "1 Component of " & lComponents & " has been identified as ""Used Common Component"" and updated because the code origin had changed (" & Left(sReplaced, Len(sReplaced) - 2) & ")."
               End Select
        Case Else
            Select Case lReplaced
                Case 0:     sMsg = lClonedRaw & " Components of " & lComponents & " had been identified as ""Used Common Components"". None had been updated since none of the code origins had changed."
                Case 1:     sMsg = lClonedRaw & " Components of " & lComponents & " had been identified as ""Used Common Components"". One had been updated since the code origin had changed (" & Left(sReplaced, Len(sReplaced) - 2) & ")."
                Case Else:  sMsg = lClonedRaw & " Components of " & lComponents & " had been identified as ""Used Common Components"". " & lReplaced & " have been updated because the code origin had changed (" & Left(sReplaced, Len(sReplaced) - 2) & ")."
            End Select
    End Select
    If Len(sMsg) > 255 Then sMsg = Left(sMsg, 251) & " ..."
    Application.StatusBar = sMsg
    
xt: mErH.EoP ErrSrc(PROC)
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

Private Function KindOfCodeChange( _
        ByVal kind_of_comp As enKindOfComp, _
        Optional ByVal raw_remote As clsComp, _
        Optional ByVal raw_cloned As clsComp) As enKindOfCodeChange
' ---------------------------------------------------------------
' Returns the kind of code change.
' - The "Raw" has changed when the Component is a used Raw
'   and its Export File differs from the Export File of the
'   Raw Component hosted in another Workbook
' - The "Used Raw Component" or the "Neither Hosted Nor Used
'   Raw Component" has changed when the code of the
'   VBComponent differs from its Export File.
' - Both, the "Raw Hosted in another Workbook" and the "Used
'   Raw Component have changed, when both of the above is
'   true for a "Used Raw Component".
' -----------------------------------------------------------
    Dim bRawRemoteChanged   As Boolean
    Dim bRawClonedChanged   As Boolean
    Dim bCodeChanged        As Boolean
    
    Select Case kind_of_comp
        Case enClonedRaw
            bRawClonedChanged = raw_cloned.CodeChanged(raw_remote:=raw_remote, raw_clone:=raw_cloned)
        Case enRemoteRaw
            bRawRemoteChanged = raw_remote.CodeChanged(raw_remote:=raw_remote, raw_clone:=raw_cloned)
        Case kind_of_comp = enNoRaw
            bCodeChanged = cComp.CodeChanged()
    End Select
    
    If bRawClonedChanged And bRawRemoteChanged Then
        KindOfCodeChange = enRawCloneAndRemote
    ElseIf bRawClonedChanged And Not bRawRemoteChanged Then
        KindOfCodeChange = enRawCloneOnly
    ElseIf Not bRawClonedChanged And bRawRemoteChanged Then
        KindOfCodeChange = enRawRemoteOnly
    ElseIf bCodeChanged Then
        KindOfCodeChange = enUsedOnly
    End If
    
End Function

