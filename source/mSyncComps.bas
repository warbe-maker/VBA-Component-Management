Attribute VB_Name = "mSyncComps"
Option Explicit
' -----------------------------------------------------
' Standard Module mSyncComps
' Provides the services to synchronize VBComponents
' between a source and a target Workbook/VB-Project.
' All public services provide two modes:
' - Confirm Collects synchronization issues for
'   confirmation and all collects all changed
'   VBComponents as
'
' Public services:
' - SyncNew
' - SyncObsolete
' - SyncCodeChanges
'
' -----------------------------------------------------
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncComps." & s
End Function

Public Sub SyncCodeChanges()
' -----------------------------------------------------
' When lMode=Confirm all components which had changed
' are collected and provided for confirmation else the
' changes are syncronized.
' -----------------------------------------------------
    Const PROC = "SyncCodeChanges"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim vbc         As VBComponent
    Dim sCaption    As String
    Dim SourceComp  As clsRaw
    Dim TargetComp  As clsComp
    
    For Each vbc In Sync.Source.VBProject.VBComponents
        If mComp.IsSheetDocMod(vbc) Then GoTo next_vbc
        Set SourceComp = New clsRaw
        Set SourceComp.Wrkbk = Sync.Source
        SourceComp.CompName = vbc.Name
        If Not SourceComp.Exists(Sync.Target) Then GoTo next_vbc
        
        Set TargetComp = New clsComp
        Set TargetComp.Wrkbk = Sync.Target
        TargetComp.CompName = vbc.Name
        SourceComp.CloneExpFileFullName = TargetComp.ExpFileFullName
        If Not SourceComp.Changed(TargetComp) Then GoTo next_vbc
        
        Stats.Count sic_non_doc_mods_code
        Log.ServicedItem = vbc
        
        If Sync.Mode = Confirm Then
            Sync.ConfInfo = "Changed"
            sCaption = "Display code changes" & vbLf & vbLf & vbc.Name & vbLf
            If Not Sync.Changed.Exists(sCaption) _
            Then Sync.Changed.Add sCaption, SourceComp
        Else
            Log.ServicedItem = vbc
            mRenew.ByImport rn_wb:=Sync.Target _
                          , rn_comp_name:=vbc.Name _
                          , rn_exp_file_full_name:=SourceComp.ExpFileFullName
            Log.Entry = "Renewed/updated by import of '" & SourceComp.ExpFileFullName & "'"
        End If
        
        Set TargetComp = Nothing
        Set SourceComp = Nothing
next_vbc:
    Next vbc

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SyncNew()
' ----------------------------------------------------
' Synchronize new components in the source Workbook
' (Sync.Source) with the target Workbook (Sync.Target).
' In lMode=Confirmation only the syncronization infos
' are collect for being confirmed.
' ----------------------------------------------------
    Const PROC = "SyncNew"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim vbc         As VBComponent
    Dim SourceComp  As clsRaw
    
    With Sync
        For Each vbc In Sync.Source.VBProject.VBComponents
            If vbc.Type = vbext_ct_Document Then GoTo next_vbc
            If vbc.Type = vbext_ct_ActiveXDesigner Then GoTo next_vbc
            
            Set SourceComp = New clsRaw
            Set SourceComp.Wrkbk = .Source
            SourceComp.CompName = vbc.Name
            If mComp.Exists(.Target, vbc.Name) Then GoTo next_vbc
            
            '~~ No component exists under the source component's name
            Log.ServicedItem = vbc
            Stats.Count sic_non_doc_mod_new
            
            If .Mode = Confirm Then
                .ConfInfo = "New! Corresponding source Workbook Export-File will by imported."
            Else
                .Target.VBProject.VBComponents.Import SourceComp.ExpFileFullName
                Log.Entry = "Component imported from Export-File '" & SourceComp.ExpFileFullName & "'"
            End If
            
            Set SourceComp = Nothing
next_vbc:
        Next vbc
    End With

xt: Set SourceComp = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SyncObsolete()
' ---------------------------------------------------------
' Synchronize obsolete components in the source Workbook
' (Sync.Source) with the target Workbook (Sync.Target). In
' lMode=Confirm only the syncronization infos are collected
' for being confirmed.
' ---------------------------------------------------------
    Const PROC = "SyncObsolete"
    
    On Error GoTo eh
    Dim vbc         As VBComponent
    Dim sType       As String
    Dim TargetComp  As clsComp
    
    With Sync
        '~~ Collect obsolete Standard Modules, Class modules, and UserForms
        For Each vbc In .Target.VBProject.VBComponents
            If vbc.Type = vbext_ct_Document Then GoTo next_vbc
            Set TargetComp = New clsComp
            Set TargetComp.Wrkbk = .Target
            TargetComp.CompName = vbc.Name
            If TargetComp.Exists(.Source) Then GoTo next_vbc
            
            Log.ServicedItem = vbc
            Stats.Count sic_non_doc_mod_obsolete
            
            If .Mode = Confirm Then
                .ConfInfo = "Obsolete!"
            Else
                sType = TargetComp.TypeString
                .Target.VBProject.VBComponents.Remove vbc
                Log.Entry = "Removed!"
            End If
            Set TargetComp = Nothing
next_vbc:
        Next vbc
    End With

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

