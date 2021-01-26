Attribute VB_Name = "mCompMan"
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Standard Module mCompMan
'          Services for the management of VBComponents in Workbooks provided:
'          - stored within the 'RootServicedByCompMan'
'          - have the Conditional Compile Argument 'CompMan = 1'
'          - have 'CompMan' referenced
'          - the Workbook resides in its own dedicated folder
'          - the Workbook calls the '' service with the Open event
'          - the Workbook calls the '' service with the Save event
' Usage:   This Workbbok's services are available as 'CompMan' AddIn and will
'          be called from the serviced Workbook as follows:
'
'           Private Sub Workbook_Open()
'           #If CompMan Then
'               On Error Resume Next
'               mCompMan.UpdateRawClones uc_wb:=ThisWorkbook _
'                                      , uc_hosted:=HOSTED_RAWS
'           #End If
'           End Sub
'
'           Private Sub Workbook_BeforeSave(...)
'           #If CompMan Then
'               mCompMan.ExportChangedComponents ec_wb:=ThisWorkbook _
'                                              , ec_hosted:=HOSTED_RAWS
'           #End If
'           End Sub
'
' ----------------------------------------------------------------------------
' Services:
' - DisplayCodeChange       Displays the current difference between a
'                           component's code and its current Export File
' - ExportAll               Exports all components into the Workbook's
'                           dedicated folder (created when not existing)
' - ExportChangedComponents Exports all components of which the code in the
'                           Export File differs from the current code.
' - Service
'
' Uses Common Components: - mBasic
'                         - mErrhndlr
'                         - mFile
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

Public cComp            As clsComp
Public cRaw             As clsRaw
Public cLog             As clsLog
Public asNoSynch()      As String
Public lMaxCompLength   As Long
Private dctHostedRaws   As Dictionary
Private sService        As String

Private Property Get HostedRaws() As Variant:        Set HostedRaws = dctHostedRaws: End Property

Private Property Let HostedRaws(ByVal hr As Variant)
' ---------------------------------------------------
' Saves the names of the hosted raw components (hr)
' to the Dictionary (dctHostedRaws).
' ---------------------------------------------------
    Dim v       As Variant
    Dim sComp   As String
    
    If dctHostedRaws Is Nothing Then
        Set dctHostedRaws = New Dictionary
    Else
        dctHostedRaws.RemoveAll
    End If
    For Each v In Split(hr, ",")
        sComp = Trim$(v)
        If Not dctHostedRaws.Exists(sComp) Then
            dctHostedRaws.Add sComp, sComp
        End If
    Next v
    
End Property

Public Property Get Service() As String:            Service = sService:             End Property

Public Property Let Service(ByVal srvc As String):  sService = srvc:                End Property

Private Function Clones( _
                  ByVal cl_wb As Workbook) As Dictionary
' ------------------------------------------------------
' Returns a Dictionary with clone component's object as
' the key and their kind of code change as item.
' ------------------------------------------------------
    Const PROC = "Clones"
    
    On Error GoTo eh
    Dim vbc As VBComponent
    Dim dct As New Dictionary
    Dim fso As New FileSystemObject
    
    mErH.BoP ErrSrc(PROC)
    If cLog Is Nothing Then
        Set cLog = New clsLog
        cLog.ServiceProvided(svp_by_wb:=ThisWorkbook, svp_for_wb:=cl_wb) = ErrSrc(PROC)
    End If
    For Each vbc In cl_wb.VBProject.VBComponents
        Set cComp = New clsComp
        With cComp
            .Wrkbk = cl_wb
            .CompName = vbc.name
            If .KindOfComp = enRawClone Then
                Set cRaw = New clsRaw
                cRaw.CompName = .CompName
                cRaw.ExpFile = fso.GetFile(FilePath:=mHostedRaws.ExpFilePath(.CompName))
                cRaw.ExpFilePath = .ExpFile.Path
                cRaw.HostFullName = mHostedRaws.HostFullName(comp_name:=.CompName)
                cRaw.CloneExpFilePath = .ExpFilePath
                If cRaw.Changed Or .Changed Then
                    dct.Add vbc, vbc.name
                End If
            End If
        End With
        Set cComp = Nothing
        Set cRaw = Nothing
    Next vbc

xt: mErH.EoP ErrSrc(PROC)
    Set Clones = dct
    Set fso = Nothing
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Private Function CodeModuleIsEmpty(ByVal vbc As VBComponent) As Boolean
    With vbc.CodeModule
        CodeModuleIsEmpty = .CountOfLines = 0 Or .CountOfLines = 1 And Len(.Lines(1, 1)) < 2
    End With
End Function

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
        sExpFileRaw = mHostedRaws.ExpFilePath(comp_name:=cmp_comp_name)
        sExpFileClone = .ExpFilePath
    
        mFile.Compare file_left_full_name:=sExpFileClone _
                    , file_right_full_name:=sExpFileRaw _
                    , file_left_title:="The cloned raw's current code in Workbook/VBProject " & cComp.WrkbkBaseName & " (" & sExpFileClone & ")" _
                    , file_right_title:="The remote raw's current code in Workbook/VBProject " & mBasic.BaseName(mHostedRaws.HostFullName(.CompName)) & " (" & sExpFileRaw & ")"

    End With
    Set cComp = Nothing

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Function CompExists( _
                     ByVal ce_wb As Workbook, _
                     ByVal ce_comp_name As String) As Boolean
' -----------------------------------------------------------
' Returns TRUE when the component (ce_comp_name) exists in
' the Workbook (ce_wb).
' -----------------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = ce_wb.VBProject.VBComponents(ce_comp_name).name
    CompExists = Err.Number = 0
End Function

Private Sub DeleteObsoleteExpFiles(ByVal do_wb As Workbook, _
                                   ByVal do_log As clsLog)
' --------------------------------------------------------------
' Delete Export Files the component does not or no longer exist.
' --------------------------------------------------------------
    Const PROC = "DeleteObsoleteExpFiles"
    
    On Error GoTo eh
    Dim cllRemove   As New Collection
    Dim sFolder     As String
    Dim fso         As New FileSystemObject
    Dim fl          As FILE
    Dim v           As Variant
    Dim cComp       As New clsComp
    Dim sComp       As String
    
    With cComp
        .Wrkbk = do_wb ' assignment provides the Workbook's dedicated Export Folder
        sFolder = .ExpFolder
    End With
    
    With fso
        '~~ Collect obsolete Export Files
        For Each fl In .GetFolder(sFolder).Files
            Select Case .GetExtensionName(fl.Path)
                Case "bas", "cls", "frm", "frx"
                    sComp = .GetBaseName(fl.Path)
                    If Not cComp.Exists(sComp) Then
                        cllRemove.Add fl.Path
                    End If
            End Select
        Next fl
    
        For Each v In cllRemove
            .DeleteFile v
            do_log.Action = "Obsolete Export File '" & v & "' deleted"
        Next v
    End With
    
xt: Set cComp = Nothing
    Set cllRemove = Nothing
    Set fso = Nothing
    Exit Sub

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
        sTempFolder = .GetFile(cComp.ExpFilePath).ParentFolder & "\Temp"
        If Not .FolderExists(sTempFolder) Then .CreateFolder sTempFolder
        sExpFileTemp = sTempFolder & "\" & cComp.CompName & cComp.Extension
        cComp.VBComp.Export sExpFileTemp
        Set flExpTemp = .GetFile(sExpFileTemp)
    End With

    With cComp
        mFile.Compare file_left_full_name:=sExpFileTemp _
                    , file_right_full_name:=.ExpFilePath _
                    , file_left_title:="The component's current code in Workbook/VBProject " & cComp.WrkbkBaseName & " ('" & sExpFileTemp & "')" _
                    , file_right_title:="The component's currently exported code in '" & .ExpFilePath & "'"

    End With
    
xt: If fso.FolderExists(sTempFolder) Then fso.DeleteFolder (sTempFolder)
    Set cComp = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

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
        If mMe.IsAddinInstnc _
        Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The Workbook (active or provided) is the CompMan Addin instance which is impossible for this operation!"
        .Wrkbk = exp_wrkbk
        For Each vbc In .Wrkbk.VBProject.VBComponents
            .CompName = vbc.name ' this assignment provides the name for the export file
            vbc.Export .ExpFilePath
        Next vbc
    End With

xt: Set cComp = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub ExportChangedComponents( _
                             ByVal ec_wb As Workbook, _
                    Optional ByVal ec_hosted As String = vbNullString)
' --------------------------------------------------------------------
' Exclusively performed/trigered by the Before_Save event:
' - Any code change (detected by the comparison of a temporary export
'   file with the current export file) is backed-up/exported
' - Outdated Export Files (components no longer existing) are removed
' - Clone code modifications update the raw code when confirmed by the
'   user
' --------------------------------------------------------------------------
    Const PROC = "ExportChangedComponents"
    
    On Error GoTo eh
    Dim lCompMaxLen         As Long
    Dim vbc                 As VBComponent
    Dim lComponents         As Long
    Dim lCompsRemaining     As Long
    Dim lExported           As Long
    Dim sExported           As String
    Dim bUpdated            As Boolean
    Dim lUpdated            As Long
    Dim sUpdated            As String
    Dim sMsg                As String
    Dim fso                 As New FileSystemObject
    Dim sServiced           As String
    Dim sProgressDots       As String
    Dim sStatus             As String
    
    mErH.BoP ErrSrc(PROC)
    '~~ Prevent any action for a Workbook opened with any irregularity
    '~~ indicated by an '(' in the active window or workbook fullname.
    If WbkIsRestoredBySystem(ec_wb) Then GoTo xt
    If Not WbkInServicedRoot(ec_wb) Then GoTo xt
    If mMe.AddInPaused Then GoTo xt
    
    mCompMan.Service = PROC & " for '" & ec_wb.name & "': "
    sStatus = mCompMan.Service
    lCompMaxLen = MaxCompLength(wb:=ec_wb)
    Set cLog = New clsLog
    cLog.ServiceProvided(svp_by_wb:=ThisWorkbook, svp_for_wb:=ec_wb, svp_new_log:=False) = ErrSrc(PROC)

    DeleteObsoleteExpFiles do_wb:=ec_wb, do_log:=cLog
    MaintainHostedRaws mh_hosted:=ec_hosted _
                     , mh_wb:=ec_wb
    
    lComponents = ec_wb.VBProject.VBComponents.Count
    lCompsRemaining = lComponents
    sProgressDots = String$(lCompsRemaining, ".")

    For Each vbc In ec_wb.VBProject.VBComponents
        If CodeModuleIsEmpty(vbc) Then GoTo next_vbc
        Set cComp = New clsComp
        sProgressDots = Left(sProgressDots, Len(sProgressDots) - 1)
        Application.StatusBar = sStatus & sProgressDots & sExported & " " & vbc.name
        mTrc.BoC ErrSrc(PROC) & " " & vbc.name
        Set cComp = New clsComp
        With cComp
            .Wrkbk = ec_wb
            .CompName = vbc.name
            sServiced = .Wrkbk.name & " Component " & .CompName & " "
            sServiced = sServiced & String(lCompMaxLen - Len(.CompName), ".")
            cLog.ServicedItem = sServiced
            If Not .Changed Then GoTo next_vbc
        End With
                
        Select Case cComp.KindOfComp
            Case enRawClone
                '~~ Establish a component class object which represents the cloned raw's remote instance
                '~~ which is hosted in another Workbook
                Set cRaw = New clsRaw
                With cRaw
                    '~~ Provide all available information rearding the remote raw component
                    '~~ Attention must be paid to the fact that the sequence of property assignments matters
                    .HostFullName = mHostedRaws.HostFullName(comp_name:=cComp.CompName)
                    .CompName = cComp.CompName
                    .ExpFile = fso.GetFile(mHostedRaws.ExpFilePath(comp_name:=.CompName))
                    .ExpFilePath = .ExpFile.Path
                    .CloneExpFilePath = cComp.ExpFilePath
                    If Not .Changed And Not cComp.Changed Then GoTo next_vbc
                End With
                
                With cComp
                    If .Changed And Not cRaw.Changed Then
                        '~~ --------------------------------------------------------------------------
                        '~~ The code change in the clone component is now in question whether it is to
                        '~~ be ignored, i.e. the change is reverted with the Workbook's next open or
                        '~~ the raw code should be updated accordingly to make the change permanent
                        '~~ for all users of the component.
                        '~~ --------------------------------------------------------------------------
                        vbc.Export .ExpFilePath
                        '~~ In case the raw had been imported manually the new check for a change will indicate no change
                        If cRaw.Changed(check_again:=True) Then GoTo next_vbc
                        .ReplaceRawWithCloneWhenConfirmed rwu_updated:=bUpdated, rwu_log:=cLog ' when confirmed in user dialog
                        If bUpdated Then
                            lUpdated = lUpdated + 1
                            sUpdated = vbc.name & ", " & sUpdated
                            cLog.Action = """Remote Raw"" has been updated with code of ""Raw Clone"""
                        End If
                        
                    ElseIf Not .Changed And cRaw.Changed Then
                        '~~ -----------------------------------------------------------------------
                        '~~ The raw had changed since the Workbook's open. This case is not handled
                        '~~ along with the Workbook's Save event but with the Workbook's Open event
                        '~~ -----------------------------------------------------------------------
                    End If
                End With
            
            Case enKindOfComp.enUnknown
                '~~ This should never be the case in is thus ignored
            
            Case Else ' enInternal, enHostedRaw
                With cComp
                    If .Changed Then
                        Application.StatusBar = sStatus & vbc.name & " Export to '" & .ExpFilePath & "'"
                        vbc.Export .ExpFilePath
                        sStatus = sStatus & vbc.name & ", "
                        cLog.Action = "Changes exported to '" & .ExpFilePath & "'"
                        lExported = lExported + 1
                        If lExported = 1 _
                        Then sExported = vbc.name _
                        Else sExported = sExported & ", " & vbc.name
                        mTrc.EoC ErrSrc(PROC) & " Backup No-Raw" & vbc.name
                        GoTo next_vbc
                    End If
                
                    If .KindOfComp = enHostedRaw Then
                        If mHostedRaws.ExpFilePath(comp_name:=.CompName) <> .ExpFilePath Then
                            mHostedRaws.ExpFilePath(comp_name:=.CompName) = .ExpFilePath
                            cLog.Action = "Component's Export File Full Name registered"
                        End If
                    End If
                End With
        End Select
                                
next_vbc:
        mTrc.EoC ErrSrc(PROC) & " " & vbc.name
        lCompsRemaining = lCompsRemaining - 1
        Set cComp = Nothing
        Set cRaw = Nothing
    Next vbc
    
    sMsg = mCompMan.Service
    Select Case lExported
        Case 0:     sMsg = sMsg & "None of the " & lComponents & " components' code changed."
        Case Else:  sMsg = sMsg & lExported & " of " & lComponents & " components' code changes : " & sExported
    End Select
    If Len(sMsg) > 255 Then sMsg = Left(sMsg, 251) & " ..."
    Application.StatusBar = sMsg
    
xt: Set dctHostedRaws = Nothing
    Set cComp = Nothing
    Set cRaw = Nothing
    Set fso = Nothing
    mErH.EoP ErrSrc(PROC)   ' End of Procedure (error call stack and execution trace)
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

    Set dctHostedRaws = New Dictionary
    HostedRaws = mh_hosted
    If HostedRaws.Count <> 0 Then
        If Not mRawHosts.Exists(raw_host_base_name:=fso.GetBaseName(mh_wb.FullName)) _
        Or mRawHosts.FullName(host_base_name:=fso.GetBaseName(mh_wb.FullName)) <> mh_wb.FullName Then
            '~~ Keep a record when this Workbook hosts one or more Raw components and not is already registered
            mRawHosts.FullName(host_base_name:=fso.GetBaseName(mh_wb.FullName)) = mh_wb.FullName
            cLog.Action = "Workbook registered as a host for at least one raw component"
        End If
    
        For Each v In HostedRaws
            '~~ Keep a record for each of the raw components hosted by this Workbook
            If Not mHostedRaws.Exists(raw_comp_name:=v) _
            Or mHostedRaws.HostFullName(comp_name:=v) <> mh_wb.FullName Then
                mHostedRaws.HostFullName(comp_name:=v) = mh_wb.FullName
                cLog.Action = "Raw component '" & v & "' hosted in this Workbook registered"
            End If
        Next v
    Else
        '~~ Remove any raws still existing and pointing to this Workbook as host
        For Each v In mHostedRaws.Components
            If mHostedRaws.HostFullName(comp_name:=v) = mh_wb.FullName Then
                mHostedRaws.Remove comp_name:=v
                cLog.Action = "Component removed from '" & mRawHosts.DAT_FILE & "'"
            End If
        Next v
        If mRawHosts.Exists(fso.GetBaseName(mh_wb.FullName)) Then
            mRawHosts.Remove (fso.GetBaseName(mh_wb.FullName))
            cLog.Action = "Workbook no longer a host for at least one raw component removed from '" & mRawHosts.DAT_FILE & "'"
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

Private Function MaxCompLength(ByVal wb As Workbook) As Long
    Dim vbc As VBComponent
    If lMaxCompLength = 0 Then
        For Each vbc In wb.VBProject.VBComponents
            MaxCompLength = mBasic.Max(MaxCompLength, Len(vbc.name))
        Next vbc
    End If
End Function

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
    fl_1 = fl_left.Path
    fl_2 = fl_right.Path
    
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

Public Sub RenewComp( _
      Optional ByVal rc_exp_file_full_name As String = vbNullString, _
      Optional ByVal rc_comp_name As String = vbNullString, _
      Optional ByVal rc_wb As Workbook = Nothing)
' --------------------------------------------------------------------
' This service renews a component by re-importing an Export File.
' When the provided Export File (rc_exp_file_full_name) does exist but
' a component name has been provided a file selection dialog is
' displayed with the possible files already filtered. When no Export
' File is selected the service terminates without notice.
' When the Workbook (rc_wb) is omitted it defaults to the
' ActiveWorkbook.
' When the ActiveWorkbook or the provided Workbook is ThisWorkbook
' the service terminates without notice
'
' Uses private component:
' - clsComp provides all required properties and the RenewByIport service
' - clsLog  provided logging services
'
' Uses Common Components:
' - mErH    Common Error Handling services
'           (may be replaced by any other!)
'
' The service must be called as follows from the concerned Workbook:
'
' Application.Run CompManDev.xlsb!mRenew.ByImport _
'                , <exp_file_full_name> _
'                , <comp_name> _
'                , <serviced_workbook_object>
'
' in case the CompManDev.xlsb is established as AddIn it can be called
' from any Workbook provided the AddIn is referenced:
'
' mCompMan.RenewComp [rc_exp_file_full_name:=....] _
'                  , [rc_comp_name:=....] _
'                  , [rc_wb:=the_serviced_workbook_object]
'
' Requires: - Reference to 'Microsoft Scripting Runtime'
'
' W. Rauschenberger Berlin, Jan 2021
' --------------------------------------------------------------------
    Const PROC = "RenewComp"

    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim cComp       As New clsComp
    Dim cLog        As New clsLog
    Dim flFile      As FILE
    Dim wbTemp      As Workbook
    Dim wbActive    As Workbook
    Dim sBaseName   As String
    
    If rc_wb Is Nothing Then Set rc_wb = ActiveWorkbook
    cComp.Wrkbk = rc_wb
    If rc_exp_file_full_name <> vbNullString Then
        If Not fso.FileExists(rc_exp_file_full_name) Then
            rc_exp_file_full_name = vbNullString ' enforces selection when the component name is also not provided
        End If
    End If
    
    If Not rc_comp_name <> vbNullString Then
        cComp.CompName = rc_comp_name
        If Not CompExists(ce_wb:=rc_wb, ce_comp_name:=rc_comp_name) Then
            If rc_exp_file_full_name <> vbNullString Then
                rc_comp_name = fso.GetBaseName(rc_exp_file_full_name)
            End If
        End If
    End If
    
    If ThisWorkbook Is rc_wb Then
        Debug.Print "The service '" & ErrSrc(PROC) & "' cannot run when ThisWorkbook is identical with the ActiveWorkbook!"
        GoTo xt
    End If
    
    If rc_exp_file_full_name = vbNullString _
    And rc_comp_name = vbNullString Then
        '~~ ---------------------------------------------
        '~~ Select the Export File for the re-new service
        '~~ of which the base name will be regared as the component to be renewed.
        '~~ --------------------------------------------------------
        If mFile.SelectFile(sel_init_path:=cComp.ExpPath _
                          , sel_filters:="*.bas,*.cls,*.frm" _
                          , sel_filter_name:="File" _
                          , sel_title:="Select the Export File for the re-new service" _
                          , sel_result:=flFile) _
        Then rc_exp_file_full_name = flFile.Path
    End If
    
    If rc_comp_name <> vbNullString _
    And rc_exp_file_full_name = vbNullString Then
        cComp.CompName = rc_comp_name
        '~~ ------------------------------------------------
        '~~ Select the component's corresponding Export File
        '~~ ------------------------------------------------
        sBaseName = fso.GetBaseName(rc_exp_file_full_name)
        '~~ Select the Export File for the re-new service
        If mFile.SelectFile(sel_init_path:=cComp.ExpPath _
                          , sel_filters:="*" & cComp.Extension _
                          , sel_filter_name:="File" _
                          , sel_title:="Select the Export File for the provided component '" & rc_comp_name & "'!" _
                          , sel_result:=flFile) _
        Then rc_exp_file_full_name = flFile.Path
    End If
    
    If rc_exp_file_full_name = vbNullString Then
        MsgBox Title:="Service '" & ErrSrc(PROC) & "' will be aborted!" _
             , Prompt:="Service '" & ErrSrc(PROC) & "' will be aborted because no " & _
                       "existing Export File has been provided!" _
             , Buttons:=vbOKOnly
        GoTo xt ' no Export File selected
    End If
    
    With cComp
        If rc_comp_name <> vbNullString Then
            If fso.GetBaseName(rc_exp_file_full_name) <> rc_comp_name Then
                MsgBox Title:="Service '" & ErrSrc(PROC) & "' will be aborted!" _
                     , Prompt:="Service '" & ErrSrc(PROC) & "' will be aborted because the " & _
                               "Export File '" & rc_exp_file_full_name & "' and the component name " & _
                               "'" & rc_comp_name & "' do not indicate the same component!" _
                     , Buttons:=vbOKOnly
                GoTo xt
            End If
            .CompName = rc_comp_name
        Else
            .CompName = fso.GetBaseName(rc_exp_file_full_name)
        End If
        
        If .Wrkbk Is ActiveWorkbook Then
            Set wbActive = ActiveWorkbook
            Set wbTemp = Workbooks.Add ' Activates a temporary Workbook
            cLog.Action = "Active Workbook de-activated by creating a temporary Workbook"
        End If
    
        cLog.ServiceProvided(svp_by_wb:=ThisWorkbook, svp_for_wb:=.Wrkbk, svp_new_log:=False) = ErrSrc(PROC)
        cLog.ServicedItem = .CompName
        
        mRenew.ByImport rn_wb:=.Wrkbk _
             , rn_comp_name:=.CompName _
             , rn_exp_file_full_name:=rc_exp_file_full_name
        cLog.Action = "Component renewed/updated by (re-)import of '" & rc_exp_file_full_name & "'"
    End With
    
xt: If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
        cLog.Action = "Temporary created Workbook closed without save"
        Set wbTemp = Nothing
        If Not ActiveWorkbook Is wbActive Then
            wbActive.Activate
            cLog.Action = "De-activated Workbook '" & wbActive.name & "' re-activated"
            Set wbActive = Nothing
        End If
    End If
    Set cComp = Nothing
    Set cLog = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub UpdateRawClones( _
                     ByVal uc_wb As Workbook, _
            Optional ByVal uc_hosted As String = vbNullString)
' ------------------------------------------------------------
' Updates a clone component with the Export File of the remote
' raw component provided the raw's code has changed.
' ------------------------------------------------------------
    Const PROC = "UpdateRawClones"
    
    On Error GoTo eh
    Dim wbActive    As Workbook
    Dim wbTemp      As Workbook
    Dim sStatus     As String
    
    mErH.BoP ErrSrc(PROC)
    
    mCompMan.Service = PROC & " for '" & uc_wb.name & "': "
    If WbkIsRestoredBySystem(uc_wb) Then GoTo xt
    If Not WbkInServicedRoot(uc_wb) Then GoTo xt
    If mMe.AddInPaused Then GoTo xt
    
    Set cLog = New clsLog
    cLog.ServiceProvided(svp_by_wb:=ThisWorkbook _
                       , svp_for_wb:=uc_wb _
                       , svp_new_log:=True _
                        ) = ErrSrc(PROC)
    
    Application.StatusBar = sStatus & "Maintain hosted raws"
    MaintainHostedRaws mh_hosted:=uc_hosted _
                     , mh_wb:=uc_wb
        
    Application.StatusBar = sStatus & "De-activate '" & uc_wb.name & "'"
    If uc_wb Is ActiveWorkbook Then
        '~~ De-activate the ActiveWorkbook by creating a temporary Workbook
        Set wbActive = uc_wb
        Set wbTemp = Workbooks.Add
    End If
    
    mUpdate.RawClones urc_wb:=uc_wb _
                    , urc_comp_max_len:=MaxCompLength(wb:=uc_wb) _
                    , urc_clones:=Clones(uc_wb) _
                    , urc_service:=sService _
                    , urc_log:=cLog

xt: If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
        Set wbTemp = Nothing
        If Not ActiveWorkbook Is wbActive Then
            wbActive.Activate
            Set wbActive = Nothing
        End If
    End If
    Set dctHostedRaws = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Function WbkGetOpen(ByVal go_wb_full_name) As Workbook
    Const PROC = "WbkGetOpen"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim sWbName As String
    
    If Not fso.FileExists(go_wb_full_name) Then GoTo xt
    sWbName = fso.GetFileName(go_wb_full_name)
    If mCompMan.WbkIsOpen(io_name:=sWbName) Then
        Set WbkGetOpen = Application.Workbooks(sWbName)
    Else
        Set WbkGetOpen = Application.Workbooks.Open(go_wb_full_name)
    End If

xt: Set fso = Nothing
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function WbkInServicedRoot(ByVal idr_wb As Workbook) As Boolean
    WbkInServicedRoot = InStr(idr_wb.Path, mMe.RootServicedByCompMan) <> 0
End Function

Public Function WbkIsOpen( _
           Optional ByVal io_name As String = vbNullString, _
           Optional ByVal io_full_name As String) As Boolean
' ------------------------------------------------------------
' When the full name is provided the check spans all Excel
' instances else only the current one.
' ------------------------------------------------------------
    Const PROC = ""
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim xlApp   As Excel.Application
    
    If io_name = vbNullString And io_full_name = vbNullString Then GoTo xt
    
    If io_full_name <> vbNullString Then
        '~~ With the full name the open test spans all application instances
        If Not fso.FileExists(io_full_name) Then GoTo xt
        If io_name = vbNullString Then io_name = fso.GetFileName(io_full_name)
        On Error Resume Next
        Set xlApp = GetObject(io_full_name).Application
        WbkIsOpen = Err.Number = 0
    Else
        On Error Resume Next
        io_name = Application.Workbooks(io_name).name
        WbkIsOpen = Err.Number = 0
    End If

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function WbkIsRestoredBySystem(ByVal rbs_wb As Workbook) As Boolean
    WbkIsRestoredBySystem = InStr(ActiveWindow.caption, "(") <> 0 _
                         Or InStr(rbs_wb.FullName, "(") <> 0
End Function
