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
    Dim SourceComp  As clsComp
    Dim TargetComp  As clsComp
    
    For Each vbc In Sync.source.VBProject.VBComponents
'        If mComp.IsSheetDocMod(vbc) Then GoTo next_vbc

        Set SourceComp = New clsComp
        With SourceComp
            Set .Wrkbk = Sync.source
            .CompName = vbc.Name
            If Not .Exists(Sync.Target) Then GoTo next_vbc
        End With
        
        Set TargetComp = New clsComp
        With TargetComp
            Set .Wrkbk = Sync.Target
            .CompName = vbc.Name
        End With
        
        If Not mService.FilesDiffer(fd_exp_file_1:=SourceComp.ExpFile _
                                  , fd_exp_file_2:=TargetComp.ExpFile) Then GoTo next_vbc
        
        Stats.Count sic_non_doc_mods_code
        Log.ServicedItem = vbc
        
        If Sync.Mode = Confirm Then
            Sync.ConfInfo = "Changed"
            sCaption = "Display code changes" & vbLf & vbLf & vbc.Name & vbLf
            If Not Sync.Changed.Exists(sCaption) _
            Then Sync.Changed.Add sCaption, SourceComp
        Else
            Log.ServicedItem = vbc
            If mComp.IsWrkbkDocMod(vbc) Or mComp.IsSheetDocMod(vbc) Then
                mSync.ByCodeLines sync_target_comp_name:=vbc.Name _
                                , wb_source_full_name:=SourceComp.Wrkbk.FullName _
                                , sync_source_codelines:=SourceComp.CodeLines
                Log.Entry = "Code updated line-by-line with code from Export-File '" & SourceComp.ExpFileFullName & "'"
            Else
                mRenew.ByImport bi_wb_serviced:=Sync.Target _
                              , bi_comp_name:=vbc.Name _
                              , bi_exp_file:=SourceComp.ExpFileFullName
                Log.Entry = "Renewed/updated by import of '" & SourceComp.ExpFileFullName & "'"
            End If
        End If
        
        Set TargetComp = Nothing
        Set SourceComp = Nothing
next_vbc:
    Next vbc

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
    Dim SourceComp  As clsComp
    
    With Sync
        For Each vbc In Sync.source.VBProject.VBComponents
            If vbc.Type = vbext_ct_Document Then GoTo next_vbc
            If vbc.Type = vbext_ct_ActiveXDesigner Then GoTo next_vbc
            
            Set SourceComp = New clsComp
            With SourceComp
                Set .Wrkbk = Sync.source
                .CompName = vbc.Name
                If .Exists(Sync.Target) Then GoTo next_vbc
            End With
            
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
            If TargetComp.Exists(.source) Then GoTo next_vbc
            
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
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

