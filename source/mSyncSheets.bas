Attribute VB_Name = "mSyncSheets"
Option Explicit

Public Sub SyncCode()
' -----------------------------------------------
' When lMode=Confirm all sheets which had changed
' are collected and provided for confirmation
' else the changes are syncronized.
' -----------------------------------------------
    Const PROC = "SyncCode"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim vbc         As VBComponent
    Dim sCaption    As String
    Dim sExpFile    As String
    Dim SourceComp  As clsRaw
    Dim TargetComp  As clsComp
    
    With Sync
        For Each vbc In .Source.VBProject.VBComponents
            If Not vbc.Type = vbext_ct_Document Then GoTo next_sheet
            If Not mComp.IsSheetDocMod(vbc) Then GoTo next_sheet
            
            Set SourceComp = New clsRaw
            Set SourceComp.Wrkbk = .Source
            SourceComp.CompName = vbc.Name
            If Not SourceComp.Exists(.Target) Then GoTo next_sheet
            
            Set TargetComp = New clsComp
            Set TargetComp.Wrkbk = .Target
            TargetComp.CompName = vbc.Name
            SourceComp.CloneExpFileFullName = TargetComp.ExpFileFullName
            If Not SourceComp.Changed(TargetComp) Then GoTo next_sheet
            
            Log.ServicedItem = vbc
            Stats.Count sic_non_doc_mods_code
            
            If .Mode = Confirm Then
                .ConfInfo = "Code changed!"
                sCaption = "Display code changes" & vbLf & vbLf & vbc.Name & vbLf
                If Not .Changed.Exists(sCaption) _
                Then .Changed.Add sCaption, SourceComp
            Else
                sExpFile = SourceComp.ExpFileFullName
                mSync.ByCodeLines sync_target_comp_name:=vbc.Name _
                                , wb_source_full_name:=SourceComp.Wrkbk.FullName _
                                , sync_source_codelines:=SourceComp.CodeLines
                Log.Entry = "Code updated line-by-line with code from Export-File '" & sExpFile & "'"
            End If
            Set SourceComp = Nothing
            Set TargetComp = Nothing
next_sheet:
        Next vbc
    End With

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SyncCodeName()
' ------------------------------------
'
' ------------------------------------
    Const PROC = "SyncCodeName"
    
    On Error GoTo eh
    Dim v                       As Variant
    Dim wsTarget                As Worksheet
    Dim vbc                     As VBComponent
    Dim sSourceSheetCodeName    As String
    Dim sSourceSheetName        As String
    Dim sTargetSheetCodeName    As String
    Dim sTargetSheetName        As String
    
    For Each v In Sync.SourceSheets
        sSourceSheetName = Sync.SourceSheets(v)
        sSourceSheetCodeName = SheetCodeName(Sync.Source, sSourceSheetName)
        If Not SheetExists(wb:=Sync.Target _
                         , sh1_name:=sSourceSheetName _
                         , sh1_code_name:=sSourceSheetCodeName _
                         , sh2_name:=sTargetSheetName _
                         , sh2_code_name:=sTargetSheetCodeName _
                          ) _
        Then GoTo next_sheet
        If sTargetSheetCodeName <> sSourceSheetCodeName And sTargetSheetName = sSourceSheetName Then
            '~~ The sheet's CodName has changed while the sheet's Name remained unchanged
            Log.ServicedItem = wsTarget
            Stats.Count sic_sheets_codename
            
            If Sync.Mode = Confirm Then
                Sync.ConfInfo = "CodeName change to '" & sSourceSheetCodeName & "'"
            Else
                For Each vbc In Sync.Target.VBProject.VBComponents
                    If vbc.Name = sTargetSheetCodeName Then
                        vbc.Name = sSourceSheetCodeName
                        '~~ When the sheet's CodeName has changed the sheet's code will only also be
                        '~~ synchronized when the CodeName is used - which should be the case because
                        '~~ there's no motivation to change it otherwise
                        Log.Entry = "CodeName changed to '" & sSourceSheetCodeName & "'"
                        Exit For
                    End If
                Next vbc
            End If
        End If
next_sheet:
    Next v

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SyncName()
' --------------------------------
'
' --------------------------------
    Const PROC = "SyncName"
    
    On Error GoTo eh
    Dim v                       As Variant
    Dim sSourceSheetCodeName    As String
    Dim sSourceSheetName        As String
    Dim sTargetSheetCodeName    As String
    Dim sTargetSheetName        As String
    
    With Sync
        For Each v In .SourceSheets
            sSourceSheetName = .SourceSheets(v)
            sSourceSheetCodeName = SheetCodeName(.Source, sSourceSheetName)
            If Not SheetExists(wb:=.Target _
                             , sh1_name:=sSourceSheetName _
                             , sh1_code_name:=sSourceSheetCodeName _
                             , sh2_name:=sTargetSheetName _
                             , sh2_code_name:=sTargetSheetCodeName _
                              ) _
            Then GoTo next_comp
            If sTargetSheetCodeName = sSourceSheetCodeName And sTargetSheetName <> sSourceSheetName Then
                Log.ServicedItem = .Source.Worksheets(sSourceSheetName)
                Stats.Count sic_sheets_name
                
                '~~ The sheet's Name has changed while the sheets CodeName remained unchanged
                If .Mode = Confirm Then
                    .ConfInfo = "Name change to '" & sSourceSheetName & "'."
                    SourceSheetNameChange sSourceSheetName, sSourceSheetCodeName, sTargetSheetName, sTargetSheetCodeName
                Else
                    .Target.Worksheets(sTargetSheetName).Name = sSourceSheetName
                    Log.Entry = "Name changed to '" & sSourceSheetName & "'."
                End If
            End If
next_comp:
        Next v
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SyncNew()
' ---------------------------------------------------------------
' Synchronize new sheets in the source Workbook (Sync.Source) with
' the target Workbook (Sync.Target).
' - When the optional new sheets counter (sync_new_count) is
'   provided, the new sheets are only counted
' - In lMode=Confirm only the syncronization infos are collect
'   for being confirmed.
' Note-1: A Worksheet is regarded new when it exists in the
'         target Workbbook under its CodeName nor its Name and
'         it is asserted that sheet's name change is restricted
'         to either Name or CodeName but never both at once.
' Note-2: This procedure is called three times
'         1. To count the sheets indicated new
'         2. To get the new sheets confirmed
'         3. To copy the new sheet from the source to the target
'            Workbook
' --------------------------------------------------------------
    Const PROC = "SyncNew"
    
    On Error GoTo eh
    Dim ws                      As Worksheet
    Dim sSourceSheetName        As String
    Dim sTargetSheetName        As String
    Dim sSourceSheetCodeName    As String
    Dim sTargetSheetCodeName    As String
    Dim v                       As Variant
    
    With Sync
        For Each v In .SourceSheets
            sSourceSheetName = .SourceSheets(v)
            sSourceSheetCodeName = SheetCodeName(.Source, sSourceSheetName)
            If Not SheetExists(wb:=.Target _
                             , sh1_name:=sSourceSheetName _
                             , sh1_code_name:=sSourceSheetCodeName _
                             , sh2_name:=sTargetSheetName _
                             , sh2_code_name:=sTargetSheetCodeName _
                             ) Then
                If NameChange(sSourceSheetName, sSourceSheetCodeName) Then GoTo next_v
        
                '~~ The sheet not exist in the target Workbook under the Name nor under the CodeName.
                Set ws = .Source.Worksheets(sSourceSheetName)
                Log.ServicedItem = ws
                Stats.Count sic_sheets_new
                
                If .Mode = Count Then
                    '~~ This is just the first call for counting the potentially new sheets
                    .CountSheetsNew
                ElseIf .Mode = Confirm Then
                    '~~ This is just the second call for the collection of the sync confirmation info
                    If .SheetsNew > 0 Or .SheetsObsolete > 0 Then
                        If Not .RestrictRenameAsserted Then
                            .Ambigous = True
                            .ConfInfo = "Ambigous new! 1)"
                            .NewSheet(sSourceSheetCodeName) = sSourceSheetName
                        Else
                            .Ambigous = False
                            .ConfInfo = "New! 2)"
                            .NewSheet(sSourceSheetCodeName) = sSourceSheetName
                        End If
                    Else
                        .ConfInfo = "New!"
                        .NewSheet(sSourceSheetCodeName) = sSourceSheetName
                    End If
                Else
                    '~~ This is the third call for getting the syncronizations done
                    '~~ The new sheet is copied to the corresponding position in the target Workbook
                    .Source.Worksheets(sSourceSheetName).Copy _
                    After:=.Target.Sheets(.Target.Worksheets.Count)
                    Log.Entry = "Copied from source Workbook."
                End If
            End If
next_v:
        Next v
    End With
       
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SyncObsolete()
' --------------------------------------------------------------------
' Remove sheets in the target (Sync.Target) which are regarded
' obsolete because they do not exist in the target Workbook
' (Sync.Target) neither under their Name nor theit CodeName.
' - When the optional obsolet sheets counter (sync_obsolete_count)
'   is provided, the obsolete sheets are only counted
' - In lMode=Confirm only the syncronization infos are collected
'   for being confirmed.
' A Worksheet is finally only regarded a obsolete when:
' A) it exists in the source Workbook neither under its CodeName nor
'    its Name a n d  it had been confirmed that name changes on sheets
'    are restricted to either or but never both at once.
' B) the number of new sheets is 0
' Note: This procedure is called three times
' 1. To count the sheets indicated obsolete
' 2. To get the removal of the obsolete sheets confirmed
' 3. To remove the obsolete sheets
' -----------------------------------------------------------------
    Const PROC = "SyncObsolete"
    
    On Error GoTo eh
    Dim ws                      As Worksheet
    Dim cSource                 As clsRaw
    Dim cTarget                 As clsComp
    Dim v                       As Variant
    Dim sTargetSheetName        As String
    Dim sTargetSheetCodeName    As String
    
    With Sync
        For Each v In .TargetSheets
            sTargetSheetName = .TargetSheets(v)
            sTargetSheetCodeName = SheetCodeName(.Target, sTargetSheetName)
            If Not SheetExists(wb:=.Source _
                             , sh1_name:=sTargetSheetName _
                             , sh1_code_name:=sTargetSheetCodeName _
                             ) Then
                If NameChange(sTargetSheetName, sTargetSheetCodeName) Then GoTo next_v
                
                '~~ Target sheet not or no longer exists in source Workbook
                '~~ neither under the Name nor under the CodeName
                Set ws = .Target.Worksheets(sTargetSheetName)
                Log.ServicedItem = ws
                Stats.Count sic_sheets_obsolete
                
                If .Mode = Count Then
                    '~~ This is just the first call for counting the potentially new sheets
                    .CountSheetsObsolete
                ElseIf .Mode = Confirm Then
                    '~~ This is just the second call for the collection of the sync confirmation info
                    If .SheetsNew > 0 Or .SheetsObsolete > 0 Then
                        If Not .RestrictRenameAsserted Then
                            .Ambigous = True
                            .ConfInfo = "Ambigous obsolete! 1)"
                        Else
                            .Ambigous = False
                            .ConfInfo = "Obsolete! 2)"
                        End If
                    Else
                        .ConfInfo = "Obsolete!"
                    End If
                Else
                    If Not .RestrictRenameAsserted Then GoTo xt
                    '~~ This is a Worksheet with no corresponding component and no corresponding sheet in the source Workbook.
                    '~~ Because it has been asserted that sheets are never renamed by Name and CodeName at once
                    '~~ this Worksheet is regarded obsolete for sure and will thus now be removed
                    For Each ws In .Target.Worksheets
                        If ws.CodeName = sTargetSheetCodeName Then
                            '~~ This is the obsolete sheet to be removed
                            Application.DisplayAlerts = False
                            ws.Delete
                            Application.DisplayAlerts = True
                            Log.Entry = "Obsolete (deleted)"
                            Exit For
                        End If
                    Next ws
                End If
                Set cTarget = Nothing
                Set cSource = Nothing
            End If
next_v:
        Next v
    End With
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SyncOrder()
' ----------------------------------------------------------------------------
' Syncronize the order of the sheets in the synch target Workbook (Sync.Target)
' to appear in the same order as in the synch source Workbook (Sync.Source).
' ----------------------------------------------------------------------------
    Const PROC = "SyncOrder"
    
    On Error GoTo eh
    Dim ws      As Worksheet
    Dim sSheet  As String
    Dim i       As Long
    
    For i = 1 To Sync.Source.Worksheets.Count
        If Sync.TargetSheets.Exists(Sync.Source.Worksheets(i).Name) Then
            Set ws = Sync.Source.Worksheets(i)
            sSheet = ws.Name
            If Sync.Target.Worksheets(i).Name <> sSheet Then
                Log.ServicedItem = ws
                If i = 1 Then
                    Sync.Target.Worksheets(sSheet).Move Before:=Sheets(i + 1)
                    Log.Entry = "Order synched!"
                Else
                    Sync.Target.Worksheets(sSheet).Move After:=Sheets(i)
                    Log.Entry = "Order synched!"
                End If
            End If
        End If
    Next i

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncSheets." & s
End Function

Private Function SheetExists( _
                       ByRef wb As Workbook, _
              Optional ByRef sh1_name As String = vbNullString, _
              Optional ByRef sh1_code_name As String = vbNullString, _
              Optional ByRef sh2_name As String = vbNullString, _
              Optional ByRef sh2_code_name As String = vbNullString) As Boolean
' -----------------------------------------------------------------------------
' Returns TRUE when the sheet either under the Name (sh1_name) or under the
' CodeName (sh1_code_name) exists in Workbook (wb).
' Returns FALSE when the sheet not exists in the Workbook (wb) under either
' name. When it exists by Name or CodeName both are returned (sh2_name,
' sh2_code_name).
' -----------------------------------------------------------------------------
    Const PROC = "SheetExists"
                             
    On Error GoTo eh
    Dim ws As Worksheet
    
    If sh1_name = vbNullString And sh1_code_name = vbNullString _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "Neither a Sheet's Name nor CodeName is provided!"
    
    For Each ws In wb.Worksheets
        If sh1_name <> vbNullString Then
            If ws.Name = sh1_name Then
                sh2_name = ws.Name
                sh2_code_name = ws.CodeName
                SheetExists = True
                Exit For
            End If
        End If
        If sh1_code_name <> vbNullString Then
            If ws.CodeName = sh1_code_name Then
                sh2_name = ws.Name
                sh2_code_name = ws.CodeName
                SheetExists = True
                Exit For
            End If
        End If
    Next ws
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function
 
Private Sub SourceSheetNameChange( _
                            ByVal sh1_name As String, _
                            ByVal sh1_code_name As String, _
                            ByVal sh2_name As String, _
                            ByVal sh2_code_name As String)
' ----------------------------------------------------------
' Registers all sheet names involved in name changes.
' ----------------------------------------------------------
    With Sync
        If Not .NameChange.Exists(sh1_name) Then .NameChange.Add sh1_name, sh1_name
        If Not .NameChange.Exists(sh1_code_name) Then .NameChange.Add sh1_code_name, sh1_code_name
        If Not .NameChange.Exists(sh2_name) Then .NameChange.Add sh2_name, sh2_name
        If Not .NameChange.Exists(sh2_code_name) Then .NameChange.Add sh2_code_name, sh2_code_name
    End With
End Sub

Private Function NameChange( _
                      ByVal sh_name As String, _
                      ByVal sh_code_name As String) As Boolean
' ------------------------------------------------------------
' Returns TRUE when either name is involved in a name change.
' ------------------------------------------------------------
    NameChange = Sync.NameChange.Exists(sh_name)
    If Not NameChange Then NameChange = Sync.NameChange.Exists(sh_code_name)
End Function

