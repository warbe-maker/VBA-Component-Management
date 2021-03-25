Attribute VB_Name = "mSyncRefs"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncRefs." & s
End Function

Public Sub SyncNew()
' --------------------------------------------
' When lMode=Confirm to be synchronized
' References are collected for being confirmed
' else References are synchronized.
' --------------------------------------------
    Const PROC = "SyncReferencesNew"
    
    On Error GoTo eh
    Dim ref As Reference
    
    For Each ref In Sync.Source.VBProject.References
        If Not RefExists(Sync.Target, ref) Then
            cLog.ServicedItem = ref
            Stats.Count sic_refs_new
            If Sync.Mode = Confirm Then
                Sync.ConfInfo = "New!"
            Else
                Sync.Target.VBProject.References.AddFromGuid ref.GUID, ref.Major, ref.Minor
                cLog.Entry = "Added"
            End If
        End If
    Next ref
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SyncObsolete()
' --------------------------------------------
' When lMode=Confirm to be synchronized
' References are collected for being confirmed
' else References are synchronized.
' --------------------------------------------
    Const PROC = "SyncReferencesObsolete"
    
    On Error GoTo eh
    Dim ref     As Reference
    Dim sRef    As String
    
    For Each ref In Sync.Target.VBProject.References
        If Not RefExists(Sync.Source, ref) Then
            cLog.ServicedItem = ref
            Stats.Count sic_refs_new
            sRef = ref.Name
            If Sync.Mode = Confirm Then
                Sync.ConfInfo = "Obsolete!"
            Else
                RefRemove ref
                cLog.Entry = "Removed!"
            End If
        End If
    Next ref

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select

End Sub

Private Function RefExists( _
                     ByRef re_wb As Workbook, _
                     ByVal re_ref As Reference) As Boolean
' --------------------------------------------------------
'
' --------------------------------------------------------
    Dim ref As Reference
    
    For Each ref In re_wb.VBProject.References
        RefExists = ref.Name = re_ref.Name
        If RefExists Then Exit Function
    Next ref

End Function

Private Sub RefRemove(ByVal rr_ref As Reference)
' -------------------------------------------------
' Removes Reference (rr_ref) from Workbook (rr_wb).
' -------------------------------------------------
    Dim ref As Reference
    
    With Sync.Target.VBProject
        For Each ref In .References
            If ref.Name = rr_ref.Name Then
                .References.Remove ref
                Exit Sub
            End If
        Next ref
    End With
    
End Sub

