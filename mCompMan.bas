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
Public Enum enKindOfComp    ' the kind of Component in the sense of CompMan
        enCommonHosted      ' The Component is a Common hosted one
        enCommonUsed        ' The Component is a Common used one
        enNotCommon         ' Neither a hosted not a used Common Component (not Common)
End Enum

Public Enum enUpdateReply
    enUpdateOriginWithUsed
    enUpdateUsedWithOrigin
    enUpdateNone
End Enum

' Distinguish the code of which Workbook is allowed to be updated
Public Enum enUpdate
        enSourceOnly
        enTargetOnly
        enSourceAndTarget
End Enum
Public Enum vbcmType
    vbext_ct_StdModule = 1          ' .bas
    vbext_ct_ClassModule = 2        ' .cls
    vbext_ct_MSForm = 3             ' .frm
    vbext_ct_ActiveXDesigner = 11   ' ??
    vbext_ct_Document = 100         ' .cls
End Enum

Public Enum enKindOfChange  ' Kind of code change
    enSourceOnly            ' Only the code of the source Component had changed
    enTargetOnly            ' Only the code of the target Component had changed
    enSourceAndTarget       ' Both, the code of the source and the target Component had changed
    enNeitherNore           ' Neither code had changed
End Enum

Public dat          As clsAppData
Public wblog        As clsAppData
Private wbTarget    As Workbook     ' The Workbook of which the VB-Components are managed
                                    ' (the Workbook which "pulls" its up to date VB-Project from wbSource)
Public asNoSynch()  As String

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompMan" & "." & sProc
End Function

Public Sub ExportAll(Optional ByVal wb As Workbook = Nothing)
' -----------------------------------------------------------
'
' -----------------------------------------------------------
    Const PROC = "ExportAll"
    
    On Error GoTo eh
    Dim cManagedComponent   As New clsComp
    Dim vbc                 As VBComponent
    
    mErH.BoP ErrSrc(PROC)
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    With cManagedComponent
        If wbAddIn.IsAddinInstance _
        Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The Workbook (active or provided) is the CompMan Addin instance which is impossible for this operation!"
        .Host = wb
        For Each vbc In .Host.VBProject.VBComponents
            .VBComp = vbc
            .BackUpCode
        Next vbc
        
    End With

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub ExportChangedComponents(ByVal wb As Workbook, _
                          Optional ByVal sHosted As String = vbNullString)
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
    Dim cMngdComp           As New clsComp
    Dim fl                  As File
    Dim sFolder             As String
    Dim flOriginExportFile  As File
    Dim sOriginHostFullName As String
    Dim lComponents         As Long
    Dim lExported           As Long
    Dim lRemoved            As Long
    Dim sExported           As String
    Dim bUpdated            As Boolean
    Dim lUpdated            As Long
    Dim sUpdated            As String
    Dim sMsg                As String

    mErH.BoP ErrSrc(PROC)
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If InStr(wb.FullName, "(") <> 0 Then GoTo xt ' This is a recovered version !
    
    Set dctComps = New Dictionary
    
    With cMngdComp
        .Host = wb
        .HostedCommonComponents = sHosted
        lCompMaxLen = mCommDat.CommCompsMaxLenght
        
        '~~ Keep a record when this Workbook hosts one or more Common Components
        If sHosted <> vbNullString Then .RegisterAsHostWorkbook
        
        For Each vbc In wb.VBProject.VBComponents
            .VBComp = vbc
            '~~ Register the Component when it is regarded Common for other VBProjects which
            '~~ is a primary precondition to get it automatically updated by means of CompMan in VBProjects using it.
            If .IsHostedCommonComponent Then .RegisterAsCommonComponent

            If .CodeModuleIsEmpty Then
                GoTo next_vbc
            End If
            lComponents = lComponents + 1
            dctComps.Add .Name, vbc
            If .IsHostedCommonComponent Then
                If Not .IsUsedCommonComponent Then
                    '~~ Register the hosted Common Component yet not registered
                    mCommDat.CommCompHostWorkbookBaseName(.Name) = .HostBaseName
                    mCommDat.CommCompExpFileFullName(.Name) = .ExportFileFullName
                End If
            End If
            
            If .CodeChanged Then
                .BackUpCode
                lExported = lExported + 1
                sExported = .Name & ", " & sExported
                If .IsCommonUsed(flOriginExportFile, sOriginHostFullName) Then
                    '~~ This is regarded an unusual code change because instead of maintaining the origin code
                    '~~ of a Common Component in ist "host" VBProject it had been changed in the using VBProject.
                    '~~ Nevertheless updating the origin code with this change is possible when explicitely confirmed.
                    .CommonOrigin.ComponentName = .ComponentName
                    With .CommonOrigin
                        .ExportFile = flOriginExportFile
                        .HostFullName = sOriginHostFullName
                    End With
                    .UpdateOriginWithUsedWhenConfirmedByUser bUpdated ' when confirmed in user dialog
                    If bUpdated Then
                        lUpdated = lUpdated + 1
                        sUpdated = .Name & ", " & sUpdated
                    End If
                End If
            End If
next_vbc:
        Next vbc
        
        '~~ Remove Export Files of no longer existent components
        Set dctRemove = New Dictionary
        sFolder = .ExportFolder.Path
        With New FileSystemObject
            For Each fl In .GetFolder(sFolder).Files
                Select Case .GetExtensionName(fl.Path)
                    Case "bas", "cls", "frm", "frx"
                        sComp = .GetBaseName(fl.Path)
                        If Not dctComps.Exists(sComp) Then
                            dctRemove.Add fl, fl.Name
                        End If
                End Select
            Next fl
        
            For Each v In dctRemove
                .DeleteFile v.Path
                lRemoved = lRemoved + 1
            Next v
        End With
    
        Select Case lExported
            Case 0:     sMsg = "None of the " & lComponents & " Components in Workbook " & .Host.Name & " had been changed/exported/backed up."
            Case 1:     sMsg = " 1 Component (of " & lComponents & ") in Workbook " & .Host.Name & " had been exported/backed up: " & left(sExported, Len(sExported) - 2)
            Case Else:  sMsg = lExported & " Components (of " & lComponents & ") in Workbook " & .Host.Name & " had been exported/backed up: " & left(sExported, Len(sExported) - 1)
        End Select
        If Len(sMsg) > 255 Then sMsg = left(sMsg, 251) & " ..."
        Application.StatusBar = sMsg
        
    End With
    
xt: mErH.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub ExportComponent( _
           Optional ByVal wb As Workbook = Nothing, _
           Optional ByVal sComp As String = vbNullString)
' -------------------------------------------------------
'
' -------------------------------------------------------
    Const PROC = "ExportComponent"
    
    On Error GoTo eh
    Dim cManagedComponent   As New clsComp
    Dim v                   As Variant
    Dim sExported           As String
    Dim lExported           As Long
    Dim lComponents         As Long
    Dim sMsg                As String
    Dim bWhenChangedOnly    As Boolean
    
    mErH.BoP ErrSrc(PROC)
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    With cManagedComponent
        
        If wbAddIn.IsAddinInstance _
        Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The Workbook (active or provided) is the CompMan Addin instance which is impossible for this operation!"
        .Host = wb
        lComponents = wb.VBProject.VBComponents.Count
        
        If sComp = vbNullString Then
            For Each v In .SelectedComponents(bWhenChangedOnly)
                .VBComp = .Host.VBProject.VBComponents(v)
                If bWhenChangedOnly Then
                    If .CodeChanged Then
                        .BackUpCode
                        sExported = .VBComp.Name & ", " & sExported
                        lExported = lExported + 1
                    End If
                Else
                    .BackUpCode
                    sExported = .VBComp.Name & ", " & sExported
                    lExported = lExported + 1
                End If
            Next v
        Else
            If .Exists(sComp) Then
                .VBComp = .Host.VBProject.VBComponents(sComp)
                .BackUpCode
                sExported = .VBComp.Name & ", " & sExported
                lExported = lExported + 1
            End If
        End If

        Select Case lExported
            Case 0:     sMsg = "None of the " & lComponents & " Components in Workbook " & .Host.Name & " had been changed/exported/backed up."
            Case 1:     sMsg = " 1 Component (of " & lComponents & ") in Workbook " & .Host.Name & " had been exported/backed up: " & left(sExported, Len(sExported) - 2)
            Case Else:  sMsg = lExported & " Components (of " & lComponents & ") in Workbook " & .Host.Name & " had been exported/backed up: " & left(sExported, Len(sExported) - 1)
        End Select
        If Len(sMsg) > 255 Then sMsg = left(sMsg, 251) & " ..."
        Application.StatusBar = sMsg
    
    End With
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub UpdateUsedCommCompsTheOriginHasChanged( _
           ByVal wbTarget As Workbook, _
  Optional ByVal sHosted As String = vbNullString)
' --------------------------------------------------
' Updates in the target Workbook (wbTarget) any used
' Common Component's code.
' Note 1: Known Common Components are those which
'         had been logged as "hosted" with the
'         ExportChangedComponents procedure per-
'         formed with the "Before_Save" event.
' Note 2: The update is performed only when con-
'         firmed in a user dialog.
' --------------------------------------------------
    Const PROC              As String = "UpdateUsedCommCompsTheOriginHasChanged"
    
    On Error GoTo eh
    Dim vbc                 As VBComponent
    Dim lCompMaxLen         As Long
    Dim cTarget             As New clsComp
    Dim flOriginExportFile  As File
    Dim sOriginHostFullName As String
    Dim lComponents         As Long
    Dim lCommonUsed         As Long
    Dim lReplaced           As Long
    Dim sReplaced           As String
    Dim sMsg                As String

    mErH.BoP ErrSrc(PROC)
    If InStr(wbTarget.FullName, "(") <> 0 Then GoTo xt ' This is a recovered version !
    
    With cTarget
        .Host = wbTarget
        .HostedCommonComponents = sHosted
        lCompMaxLen = mCommDat.CommCompsMaxLenght
    
        For Each vbc In wbTarget.VBProject.VBComponents
            .VBComp = vbc
            Debug.Print .Name
            lComponents = lComponents + 1
            If .IsUsedCommonComponent Then
                If .IsCommonUsed(flOriginExportFile, sOriginHostFullName) Then
                    lCommonUsed = lCommonUsed + 1
                    With .CommonOrigin
                        .ExportFile = flOriginExportFile
                        .HostFullName = sOriginHostFullName
                    End With
                    If .CommonOriginCodeChanged Then
                        .UpdateUsedWithOrigin
                        lReplaced = lReplaced + 1
                        sReplaced = .Name & ", " & sReplaced
                        '~~ Register the update being used to identify a potentially relevant
                        '~~ change of the origin code
                        .CodeVersionAsOfDate = .ExportFile.DateLastModified
                    End If
                End If
            End If
        Next vbc
        
    End With
    
    Select Case lCommonUsed
        Case 0: sMsg = "No Components of " & lComponents & " had been identified as ""Common Used ""."
        Case 1
            Select Case lReplaced
                Case 0:     sMsg = "1 Component of " & lComponents & " has been identified as ""Used Common Component "" but not updated since the code origin had not changed."
                Case 1:     sMsg = "1 Component of " & lComponents & " has been identified as ""Used Common Component"" and updated because the code origin had changed (" & left(sReplaced, Len(sReplaced) - 2) & ")."
               End Select
        Case Else
            Select Case lReplaced
                Case 0:     sMsg = lCommonUsed & " Components of " & lComponents & " had been identified as ""Used Common Components"". None had been updated since none of the code origins had changed."
                Case 1:     sMsg = lCommonUsed & " Components of " & lComponents & " had been identified as ""Used Common Components"". One had been updated since the code origin had changed (" & left(sReplaced, Len(sReplaced) - 2) & ")."
                Case Else:  sMsg = lCommonUsed & " Components of " & lComponents & " had been identified as ""Used Common Components"". " & lReplaced & " have been updated because the code origin had changed (" & left(sReplaced, Len(sReplaced) - 2) & ")."
            End Select
    End Select
    If Len(sMsg) > 255 Then sMsg = left(sMsg, 251) & " ..."
    Application.StatusBar = sMsg
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Version(ByVal Version As clsAddinVersion)
' ---------------------------------------------------------------------------------------------------------------------
' Called by the development instance via Application.Run. Because the version value cannot be returned to the call via
' a ByRef argument, a class object is used instead.
' See: http://www.tushar-mehta.com/publish_train/xl_vba_cases/1022_ByRef_Argument_with_the_Application_Run_method.shtml
' ---------------------------------------------------------------------------------------------------------------------
    Version.Version = wbAddIn.AddInVersion
End Sub

