Attribute VB_Name = "mSyncComps"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncNames." & s
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
    Dim cSource     As clsRaw
    Dim cTarget     As clsComp
    
    For Each vbc In Sync.Source.VBProject.VBComponents
        If IsSheetComp(vbc) Then GoTo next_vbc
        Set cSource = New clsRaw
        Set cSource.Wrkbk = Sync.Source
        cSource.CompName = vbc.Name
        If Not cSource.Exists(Sync.Target) Then GoTo next_vbc
        
        Set cTarget = New clsComp
        Set cTarget.Wrkbk = Sync.Target
        cTarget.CompName = vbc.Name
        cSource.CloneExpFileFullName = cTarget.ExpFileFullName
        If Not cSource.Changed Then GoTo next_vbc
        
        Stats.Count sic_non_doc_mods_code
        cLog.ServicedItem = vbc
        
        If Sync.Mode = Confirm Then
            Sync.ConfInfo = "Changed"
            sCaption = "Display changes" & vbLf & "of" & vbLf & vbLf & vbc.Name & vbLf
            If Not Sync.Changed.Exists(sCaption) _
            Then Sync.Changed.Add sCaption, cSource
        Else
            cLog.ServicedItem = vbc
            mRenew.ByImport rn_wb:=Sync.Target _
                          , rn_comp_name:=vbc.Name _
                          , rn_exp_file_full_name:=cSource.ExpFileFullName
            cLog.Entry = "Renewed/updated by import of '" & cSource.ExpFileFullName & "'"
        End If
        
        Set cTarget = Nothing
        Set cSource = Nothing
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
    Dim fso     As New FileSystemObject
    Dim vbc     As VBComponent
    Dim cComp   As clsComp
    Dim cSource As clsRaw
    
    With Sync
        For Each vbc In Sync.Source.VBProject.VBComponents
            If vbc.Type = vbext_ct_Document Then GoTo next_vbc
            If vbc.Type = vbext_ct_ActiveXDesigner Then GoTo next_vbc
            
            Set cSource = New clsRaw
            Set cSource.Wrkbk = .Source
            cSource.CompName = vbc.Name
            If CompExists(.Target, vbc.Name) Then GoTo next_vbc
            
            '~~ No component exists under the source component's name
            cLog.ServicedItem = vbc
            Stats.Count sic_non_doc_mod_new
            
            If .Mode = Confirm Then
                .ConfInfo = "New! Corresponding source Workbook Export-File will by imported."
            Else
                .Target.VBProject.VBComponents.Import cSource.ExpFileFullName
                cLog.Entry = "Component imported from Export-File '" & cSource.ExpFileFullName & "'"
            End If
            
            Set cSource = Nothing
next_vbc:
        Next vbc
    End With

xt: Set cComp = Nothing
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
    Dim cTarget     As clsComp
    
    With Sync
        '~~ Collect obsolete Standard Modules, Class modules, and UserForms
        For Each vbc In .Target.VBProject.VBComponents
            If vbc.Type = vbext_ct_Document Then GoTo next_vbc
            Set cTarget = New clsComp
            Set cTarget.Wrkbk = .Target
            cTarget.CompName = vbc.Name
            If cTarget.Exists(.Source) Then GoTo next_vbc
            
            cLog.ServicedItem = vbc
            Stats.Count sic_non_doc_mod_obsolete
            
            If .Mode = Confirm Then
                .ConfInfo = "Obsolete!"
            Else
                sType = cTarget.TypeString
                .Target.VBProject.VBComponents.Remove vbc
                cLog.Entry = "Removed!"
            End If
            Set cTarget = Nothing
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

