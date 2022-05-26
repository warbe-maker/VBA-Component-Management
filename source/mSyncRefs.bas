Attribute VB_Name = "mSyncRefs"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncRefs." & s
End Function

Public Function RefExists(ByRef re_wb As Workbook, _
                          ByVal re_ref As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Reference (re_ref) exists in the VB-Project of the
' Workbook (re_wb).
' ------------------------------------------------------------------------------
    Dim ref As Reference
    
    For Each ref In re_wb.VBProject.References
        If TypeName(re_ref) = "Reference" Then
            RefExists = ref.Name = re_ref.Name
        ElseIf TypeName(re_ref) = "String" Then
            RefExists = ref.Name = re_ref Or ref.Description = re_ref
        End If
        If RefExists Then Exit Function
    Next ref

End Function

Public Sub RefRemove(ByVal rr_ref As Reference, _
            Optional ByVal rr_wb_target As Workbook = Nothing)
' ------------------------------------------------------------------------------
' Removes the Reference (rr_ref) from Workbook (rr_wb).
' ------------------------------------------------------------------------------
    Dim ref         As Reference
    Dim WbTarget    As Workbook
    
    If rr_wb_target Is Nothing Then Set WbTarget = Sync.TargetWb Else Set WbTarget = rr_wb_target
    
    With WbTarget.VBProject
        For Each ref In .References
            If ref.Name = rr_ref.Name Then
                .References.Remove ref
                Exit Sub
            End If
        Next ref
    End With
    
End Sub

Public Sub RunRefAdd(ByVal ra_ref As Reference, _
                     ByVal ra_wb_target As Workbook)
    
    If Stats Is Nothing Then Set Stats = New clsStats
    If Log Is Nothing Then
        Set Log = New clsLog
        Log.Service(new_log:=True) = SERVICE_SYNCHRONIZE
        EstablishTraceLogFile ra_wb_target
    End If
    
    ra_wb_target.VBProject.References.AddFromGuid ra_ref.GUID, ra_ref.Major, ra_ref.Minor
    
    '~~ Re-display to be synchronized items
    Application.Run ThisWorkbook.Name & "!mCompMan." & SERVICE_SYNCHRONIZE, ra_wb_target
End Sub

Public Sub RunRefRemove(ByVal rr_ref As Reference, _
                        ByVal rr_wb_target As Workbook)
' ------------------------------------------------------------------------------
' Removes the Reference (rr_ref) from Workbook (rr_wb).
' ------------------------------------------------------------------------------
    Dim ref         As Reference
        
    If Stats Is Nothing Then Set Stats = New clsStats
    If Log Is Nothing Then
        Set Log = New clsLog
        Log.Service(new_log:=True) = SERVICE_SYNCHRONIZE
        EstablishTraceLogFile rr_wb_target
    End If
    
    With rr_wb_target.VBProject
        For Each ref In .References
            If ref.Name = rr_ref.Name Then
                .References.Remove ref
                Exit Sub
            End If
        Next ref
    End With
    
    '~~ Re-display to be synchronized items
    Application.Run ThisWorkbook.Name & "!mCompMan." & SERVICE_SYNCHRONIZE, rr_wb_target
    
End Sub

Public Sub SyncNew(Optional ByVal sn_wb_target As Workbook, _
                   Optional ByVal sn_wb_source As Workbook, _
                   Optional ByVal sn_app_run As Collection = Nothing)
' ------------------------------------------------------------------------------
' When lMode=Confirm The be synchronized new References are collected for being
' confirmed, else new References are synchronized.
' ------------------------------------------------------------------------------
    Const PROC = "SyncReferencesNew"
    
    On Error GoTo eh
    Dim ref         As Reference
    Dim cll         As Collection
    Dim wbSource    As Workbook
    Dim WbTarget    As Workbook
    
    If sn_wb_target Is Nothing Then Set WbTarget = Sync.TargetWb Else Set WbTarget = sn_wb_target
    If sn_wb_source Is Nothing Then Set wbSource = Sync.SourceWb Else Set wbSource = sn_wb_source
    
    For Each ref In wbSource.VBProject.References
        If Not RefExists(WbTarget, ref) Then
            Log.ServicedItem = ref
            Stats.Count sic_refs_new
            If Not sn_app_run Is Nothing Then
                Set cll = New Collection
                cll.Add ref.Description & vbLf & vbLf & "Add"  ' The button caption
                cll.Add ThisWorkbook                    ' The servicing Workbook
                cll.Add "mSyncRefs.RunRefAdd"            ' The service to run
                cll.Add ref                             ' The Reference to add
                cll.Add WbTarget                        ' The 'Synchronization-Target-Workbook'
                sn_app_run.Add cll
            ElseIf Sync.Mode = Confirm Then
                Sync.ConfInfo = "New!"
            Else
                WbTarget.VBProject.References.AddFromGuid ref.GUID, ref.Major, ref.Minor
                Log.Entry = "Added"
            End If
        End If
    Next ref
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SyncObsolete(Optional ByVal so_wb_target As Workbook, _
                        Optional ByVal so_wb_source As Workbook, _
                        Optional ByRef so_app_run As Collection = Nothing)
' ------------------------------------------------------------------------------
' When lMode = Confirm to be synchronized obsolete References are collected for
' being confirmed else obsolete References are synchronized/removed. When a
' Collection is provided (so_app_run) Application.Run arguments are collected
' as follows: 1. The buttons caption
'             2. The servicing Workbook (so_wb_servicing)
'             3. The service to run (mSyncRef.RefRemove)
'             4. The Reference objec t to be removed
'             5. The 'Synchronization-Target-Workbook' (so_wb_target)
' ------------------------------------------------------------------------------
    Const PROC = "SyncReferencesObsolete"
    
    On Error GoTo eh
    Dim ref         As Reference
    Dim sRef        As String
    Dim WbTarget    As Workbook
    Dim wbSource    As Workbook
    Dim cll         As Collection
    
    If so_wb_target Is Nothing Then Set WbTarget = Sync.TargetWb Else Set WbTarget = so_wb_target
    If so_wb_source Is Nothing Then Set wbSource = Sync.SourceWb Else Set wbSource = so_wb_source
    
    For Each ref In WbTarget.VBProject.References
        If Not RefExists(wbSource, ref) Then
            Log.ServicedItem = ref
            Stats.Count sic_refs_new
            sRef = ref.Name
            
            If Not so_app_run Is Nothing Then
                Set cll = New Collection
                cll.Add ref.Description & vbLf & vbLf & "Remove"   ' The button caption
                cll.Add ThisWorkbook                    ' The servicing Workbook
                cll.Add "mSyncRefs.RunRefRemove"         ' The service to run
                cll.Add ref                             ' The Reference to remove
                cll.Add so_wb_target                    ' The 'Synchronization-Target-Workbook'
                so_app_run.Add cll
            ElseIf Sync.Mode = Confirm Then
                Sync.ConfInfo = "Obsolete!"
            Else
                RefRemove rr_ref:=ref, rr_wb_target:=WbTarget
                Log.Entry = "Removed!"
            End If
        End If
    Next ref

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function SyncRefs(ByVal sync_wb_target As Workbook, _
                         ByVal sync_wb_source As Workbook) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when there are References for being synchronized, which means a
' modeless dialog had been displayed in order to get them synchronized by
' clicking the corresponding button.
' ------------------------------------------------------------------------------
    Const PROC = "SyncRefs"
    
    On Error GoTo eh
    Dim cllObsolete As New Collection
    Dim cllNew      As New Collection
    Dim fSync       As fMsg
    Dim Msg         As TypeMsg
    Dim cllButtons  As Collection
    Dim MsgTitle    As String
    Dim AppRunArgs  As New Dictionary
    Dim cll         As Collection
    Dim i           As Long
    
    If Stats Is Nothing Then Set Stats = New clsStats
    If Log Is Nothing Then
        Set Log = New clsLog
        Log.Service(new_log:=True) = SERVICE_SYNCHRONIZE
        EstablishTraceLogFile sync_wb_target
    End If
    
    SyncObsolete sync_wb_target, sync_wb_source, cllObsolete
    SyncNew sync_wb_target, sync_wb_source, cllNew
    
    If cllObsolete.Count + cllNew.Count <> 0 Then
        SyncRefs = True ' indicate that at least on item to be synchronized had been fisplayed
        '~~ There's at least one Reference still to be synchronized
        MsgTitle = "Synchronization of References"
        Set fSync = mMsg.MsgInstance(MsgTitle)
        With Msg.Section(1)
            .Label.Text = "Remove:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Removes the obsolete Reference from the 'Synchronization-Target-Workbook's' VB-Project"
        End With
        With Msg.Section(2)
            .Label.Text = "Add:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Adds the new Reference to the 'Synchronization-Target-Workbook's' VB-Project"
        End With
        With Msg.Section(3)
            .Label.Text = "About:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Synchronizing References is one of the tasks of the CompMan's Synchronization service. " & _
                         "This dialog will be re-displayed until there are no more References to be synchronized. " & _
                         "The one-by-one approach of the service provides full control over what is synchronized."
        End With
        
        For i = 1 To cllObsolete.Count
            Set cll = cllObsolete(i)
            Set cllButtons = mMsg.Buttons(cllButtons, cll(1))
            mMsg.ButtonAppRun AppRunArgs, cll(1), cll(2), cll(3), cll(4), cll(5)
        Next i
        
        For i = 1 To cllNew.Count
            Set cll = cllNew(i)
            Set cllButtons = mMsg.Buttons(cllButtons, cll(1))
            mMsg.ButtonAppRun AppRunArgs, cll(1), cll(2), cll(3), cll(4), cll(5)
        Next i
        
        mMsg.Dsply dsply_title:=MsgTitle _
                 , dsply_msg:=Msg _
                 , dsply_buttons:=cllButtons _
                 , dsply_modeless:=True _
                 , dsply_buttons_app_run:=AppRunArgs
    End If
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

