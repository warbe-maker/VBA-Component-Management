Attribute VB_Name = "mSync"
Option Explicit

Private lMaxLenComp             As Long
Private lMaxLenType             As Long
Private dctChanged              As Dictionary   ' Confirm buttons and clsRaw items to display changed
Private bSyncDenied             As Boolean      ' when True the synchronization is not performed
Private bAmbigous               As Boolean      ' when True sync is only done when the below is confirmed True
Private RestrictRenameAsserted  As Boolean      ' when False a sheet's CodeName a n d its Name may be changed at once
Private cSource                 As clsRaw
Private cTarget                 As clsComp
Private SyncTarget              As Workbook
Private SyncSource              As Workbook
Private SourceSheetComps        As Dictionary
Private TargetSheetComps        As Dictionary
Private SourceSheets            As Dictionary
Private TargetSheets            As Dictionary
Private SourceSheetShapes       As Dictionary
Private TargetSheetShapes       As Dictionary
Private lSheetsObsolete         As Long
Private lSheetsNew              As Long

Public Sub ByCodeLines( _
                 ByVal sync_target_comp_name As String, _
                 ByVal sync_source_wb_full_name As String, _
        Optional ByRef sync_source_codelines As Dictionary = Nothing)
' -------------------------------------------------------------------------
' Synchronizes
'  the component (sync_target_comp_name) in the target Workbook
'  (SyncTarget) with the code (sync_source_codelines) in the Export-File
'  of the corresponding source Workbook's (sync_source_wb_full_name)
'  component
' line by line.
' When the source code lines () are not provided they are obtained from the
' source Workbook's () corresponding Export-File.
' ----------------------------------------------------------------
    Const PROC = "ByCodeLines"

    On Error GoTo eh
    Dim i       As Long: i = 1
    Dim v       As Variant
    Dim ws      As Worksheet
    Dim wbRaw   As Workbook
    
    If sync_source_codelines Is Nothing Then
        '~~ Obtain non provided code lines for the line by line syncronization
        Set wbRaw = WbkGetOpen(sync_source_wb_full_name)
        Set cSource.Wrkbk = wbRaw
        cSource.CompName = sync_target_comp_name
        Set sync_source_codelines = cSource.CodeLines
    End If
    
    With SyncTarget.VBProject.VBComponents(sync_target_comp_name).CodeModule
        If .CountOfLines > 0 _
        Then .DeleteLines 1, .CountOfLines   ' Remove all lines from the cloned raw component
        
        For Each v In sync_source_codelines    ' Insert the raw component's code lines
            Debug.Print sync_source_codelines(v)
            .InsertLines i, sync_source_codelines(v)
            i = i + 1
        Next v
    End With
                
xt: Set cSource = Nothing
    Set wbRaw = Nothing
    Set ws = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Function Center(ByVal s1 As String, _
                       ByVal l As Long, _
               Optional ByVal sFill As String = " ") As String
' ------------------------------------------------------------
' Returns s1 centered in a string with length l.
' ------------------------------------------------------------
    Dim lSpace As Long
    lSpace = Max(1, ((l - Len(s1)) / 2))
    Center = VBA.String$(lSpace, sFill) & s1 & VBA.String$(lSpace, sFill)
    Center = Right(Center, l)
End Function

Private Function CompMaxLen() As Long
' -------------------------------------------------------------
' Returns the length of the longest element which may be
' displayed with the syncronization confirmation info.
' -------------------------------------------------------------
    
    Dim vbc As VBComponent
    Dim ref As Reference
    Dim l   As Long
    Dim ws  As Worksheet
    Dim shp As Shape
    
    For Each vbc In SyncTarget.VBProject.VBComponents:  l = mBasic.Max(l, Len(vbc.Name)):   Next vbc
    For Each vbc In SyncSource.VBProject.VBComponents:  l = mBasic.Max(l, Len(vbc.Name)):   Next vbc
    For Each ref In SyncTarget.VBProject.References:    l = mBasic.Max(l, Len(ref.Name)):   Next ref
    For Each ref In SyncSource.VBProject.References:    l = mBasic.Max(l, Len(ref.Name)):   Next ref
    For Each ws In SyncSource.Worksheets
        For Each shp In ws.Shapes
            l = mBasic.Max(l, Len(shp.Name))
        Next shp
    Next ws
    For Each ws In SyncTarget.Worksheets
        For Each shp In ws.Shapes
            l = mBasic.Max(l, Len(shp.Name))
        Next shp
    Next ws
    
    
    CompMaxLen = l
End Function

Private Sub CompsChanged( _
          Optional ByRef sync_confirm_info As clsLog = Nothing)
' -----------------------------------------------------
' When syncs is provided, collect all components which
' had changed. I.e. their raw Export-File differs from
' the Clone Export-File
' else perform the syncronization.
' -----------------------------------------------------
    Const PROC = "CompsNew"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim vbc1        As VBComponent
    Dim vbc2        As VBComponent
    
    For Each vbc1 In SyncSource.VBProject.VBComponents
        
        If IsSheetComp(vbc1) Then GoTo next_vbc1
        Set cSource = New clsRaw
        Set cSource.Wrkbk = SyncSource
        cSource.CompName = vbc1.Name
        If Not cSource.Exists(SyncTarget) Then GoTo next_vbc1
        
        Set cTarget = New clsComp
        Set cTarget.Wrkbk = SyncTarget
        cTarget.CompName = vbc1.Name
        cSource.CloneExpFileFullName = cTarget.ExpFileFullName
        If Not cSource.Changed Then GoTo next_vbc1
        
        If Not sync_confirm_info Is Nothing Then
            sync_confirm_info.ServicedItem = vbc1.Name
            sync_confirm_info.Info = cTarget.TypeString & " changed"
            If Not dctChanged.Exists("Display" & vbLf & vbc1.Name & vbLf & "changes") _
            Then dctChanged.Add "Display" & vbLf & vbc1.Name & vbLf & "changes", cSource
        Else
            mRenew.ByImport rn_wb:=SyncTarget _
                          , rn_comp_name:=vbc1.Name _
                          , rn_exp_file_full_name:=cRaw.ExpFileFullName
            cLog.Entry = "Renewed by import of '" & cSource.ExpFileFullName & "'"
        End If
        
        Set cTarget = Nothing
        Set cSource = Nothing
next_vbc1:
    Next vbc1

    '~~ Collect changed Document Modules representing Worksheets
    For Each vbc2 In SyncSource.VBProject.VBComponents
        If Not vbc2.Type = vbext_ct_Document Then GoTo next_sheet
        If Not IsSheetComp(vbc2) Then GoTo next_sheet
        
        Set cSource = New clsRaw
        Set cSource.Wrkbk = SyncSource
        cSource.CompName = vbc2.Name
        If Not cSource.Exists(SyncTarget) Then GoTo next_sheet
        
        Set cTarget = New clsComp
        Set cTarget.Wrkbk = SyncTarget
        cTarget.CompName = vbc2.Name
        cSource.CloneExpFileFullName = cTarget.ExpFileFullName
        If Not cSource.Changed Then GoTo next_sheet
        
        If Not sync_confirm_info Is Nothing Then
            sync_confirm_info.ServicedItem = vbc2.Name
            sync_confirm_info.Info = "Worksheet changed (about to be updated with code from Export-File line-by-line)"
            If Not dctChanged.Exists("Display" & vbLf & vbc2.Name & vbLf & "changes") _
            Then dctChanged.Add "Display" & vbLf & vbc2.Name & vbLf & "changes", cSource
        Else
            cLog.ServicedItem = vbc2.Name
            mSync.ByCodeLines sync_target_comp_name:=vbc2.Name _
                            , sync_source_wb_full_name:=cRaw.Wrkbk.FullName _
                            , sync_source_codelines:=cRaw.CodeLines
            cLog.Entry = "Worksheet code updated line-by-line from Export-File '" & cSource.ExpFileFullName & "'"
        End If
        Set cSource = Nothing
        Set cTarget = Nothing
next_sheet:
    Next vbc2

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Function CompsNew( _
           Optional ByRef sync_confirm_info As clsLog = Nothing)
' --------------------------------------------------------------
' Synchronize new components in the source Workbook
' (SyncSource) with the target Workbook (SyncTarget).
' When the optional confirmation info (sync_confirm_info) is
' provided only the syncronization infos are collect for being
' confirmed.
' --------------------------------------------------------------
    Const PROC = "CompsNew"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim vbc     As VBComponent
    Dim cComp   As clsComp
    
    For Each vbc In SyncSource.VBProject.VBComponents
        If vbc.Type = vbext_ct_Document Then GoTo next_vbc
        If vbc.Type = vbext_ct_ActiveXDesigner Then GoTo next_vbc
        
        Set cSource = New clsRaw
        Set cSource.Wrkbk = SyncSource
        cSource.CompName = vbc.Name
        If CompExists(SyncTarget, vbc.Name) Then GoTo next_vbc
        
        '~~ No component exists under the source component's name
        If Not sync_confirm_info Is Nothing Then
            sync_confirm_info.ServicedItem = vbc.Name
            sync_confirm_info.Info = "New " & cSource.TypeString
        Else
            cLog.ServicedItem = vbc.Name
            SyncTarget.VBProject.VBComponents.Import cSource.ExpFileFullName
        End If
        
        Set cSource = Nothing
next_vbc:
    Next vbc

xt: Set cComp = Nothing
    Set fso = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function CompsObsolete( _
                Optional ByRef sync_confirm_info As clsLog = Nothing)
' -------------------------------------------------------------------
' Synchronize obsolete components in the source Workbook
' (SyncSource) with the target Workbook (SyncTarget). When
' the optional confirmation info (sync_confirm_info) is provided only
' the syncronization infos are collect for being confirmed.
' -------------------------------------------------------------------
    Const PROC = "CompsObsolete"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim lType       As vbcmType
    Dim vbc         As VBComponent
    Dim sType       As String
    Dim cTarget     As clsComp
    
    '~~ Collect obsolete Standard Modules, Class modules, and UserForms
    For Each vbc In SyncTarget.VBProject.VBComponents
        If vbc.Type = vbext_ct_Document Then GoTo next_vbc
        Set cTarget = New clsComp
        Set cTarget.Wrkbk = SyncTarget
        cTarget.CompName = vbc.Name
        If cTarget.Exists(SyncSource) Then GoTo next_vbc
        
        If Not sync_confirm_info Is Nothing Then
            sync_confirm_info.ServicedItem = vbc.Name
            sync_confirm_info.Info = "Obsolete " & cTarget.TypeString & " (will be removed)"
        Else
            cLog.ServicedItem = vbc.Name
            sType = cTarget.TypeString
            SyncTarget.VBProject.VBComponents.Remove vbc
            cLog.Entry = "Obsolete " & sType & " removed"
        End If
        Set cTarget = Nothing
next_vbc:
    Next vbc

xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

'Private Sub CountSyncSheets( _
'                      ByRef SourceSheets As Dictionary, _
'                      ByRef TargetSheets As Dictionary, _
'             Optional ByRef sync_new_count As Variant = Nothing, _
'             Optional ByRef sync_obsolete_count As Variant = Nothing)
'' -------------------------------------------------------------------
'' Count the potentially new and obsolete sheets
'' -------------------------------------------------------------------
'    SheetsNew sync_new_count:=sync_new_count
'    SheetsObsolete sync_obsolete_count:=sync_obsolete_count
'
'End Sub

Private Sub FormShapes( _
              Optional ByVal ctrl_sheet_name As String = vbNullString, _
              Optional ByVal ctrl_sheet_codename As String = vbNullString)
' ----------------------------------------------------------------------
'
' ----------------------------------------------------------------------

'    FormShapesNew
'    FormShapesObsolete
    
End Sub

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSync." & s
End Function

Private Function IsSheetComp(ByRef vbc As VBComponent) As Boolean
    IsSheetComp = vbc.Type = vbext_ct_Document And Not IsWrkbkComp(vbc)
End Function

Private Function IsWrkbkComp(ByRef vbc As VBComponent) As Boolean
    
    Dim bSigned As Boolean
    On Error Resume Next
    bSigned = vbc.Properties("VBASigned").Value
    IsWrkbkComp = Err.Number = 0
    
End Function

Private Function MaxLen( _
                  ByRef max_len_comp As Long, _
                  ByRef max_len_type As Long) As Long
    Dim vbc As VBComponent
    For Each vbc In SyncSource.VBProject.VBComponents
        max_len_comp = mBasic.Max(max_len_comp, Len(vbc.Name))
    Next vbc
    For Each vbc In SyncTarget.VBProject.VBComponents
        max_len_comp = mBasic.Max(max_len_comp, Len(vbc.Name))
    Next vbc
        
    max_len_type = mBasic.Max(max_len_type, Len("ActiveX Designer"))
    max_len_type = mBasic.Max(max_len_type, Len("Class Module"))
    max_len_type = mBasic.Max(max_len_type, Len("Document Module"))
    max_len_type = mBasic.Max(max_len_type, Len("UserForm"))
    max_len_type = mBasic.Max(max_len_type, Len("Standard Module"))
End Function

Private Function NameExists( _
                      ByRef ne_wb As Workbook, _
                      ByVal ne_nm As Name) As Boolean
    Dim nm As Name
    For Each nm In ne_wb.Names
        NameExists = nm.Name = ne_nm.Name
        If NameExists Then Exit For
    Next nm
End Function

Public Sub Names( _
  Optional ByVal wb_sheet_name As String = vbNullString, _
  Optional ByVal wb_sheet_codename As String = vbNullString)
' ----------------------------------------------------------------
' Synchronize the names in Worksheet (SyncTarget) with those in
' Workbook (SyncSource) - when either a sheet's name (wb_sheet_name) or
' a sheet's CodeName (wb_sheet_codename) is provided only those
' names which refer to that sheet.
' - Note: Obsolete names are removed but missing names cannot be
'   added but are reported in the log file Missing names must be
'   added manually in concert with the corresponding sheet design
'   changes. As a consequence, design changes should preferrably
'   be made prior copy of the Workbook for a VB-Project
'   modification.
' - Precondition: The Worksheet's CodeName and Name are identical
'   both Workbooks. I.e. these need to be synched first.
' ---------------------------------------------------------------
    Const PROC = "Names"
    
    On Error GoTo eh
    Dim nm          As Name
    Dim shRaw       As Worksheet
    Dim shClone     As Worksheet
    Dim sSheetRaw   As String
    Dim sSheetClone As String
        
    sSheetRaw = shRaw.Name
    sSheetClone = shClone.Name
    
    If wb_sheet_name <> vbNullString Then
        For Each nm In SyncTarget
            If Not NameExists(SyncSource, nm) And InStr(nm.Name, wb_sheet_name & "!") <> 0 Then
                SyncTarget.Names(nm.Name).Delete
                cLog.Entry = "Obsolete range name '" & nm.Name & "' removed"
            End If
        Next nm
    Else
        For Each nm In SyncTarget.Names
            If Not NameExists(SyncSource, nm) Then
                SyncTarget.Names(nm.Name).Delete
                cLog.Entry = "Obsolete range name '" & nm.Name & "' removed"
            End If
        Next nm
    End If

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub RefAdd(ByRef ra_wb As Workbook, _
                   ByVal ra_ref As Reference)
' ----------------------------------------------------
    ra_wb.VBProject.References.AddFromFile ra_ref.Name
End Sub

Private Sub References( _
        Optional ByRef sync_confirm_info As clsLog = Nothing)
' ---------------------------------------------------------
' Collect to be synchronized References (when sync_confirm_info Not
' is Nothing) else synchronizes the References.
' ---------------------------------------------------------

    Dim ref As Reference
    
    '~~ Add missing to clone Workbook
    For Each ref In SyncSource.VBProject.References
        If Not RefExists(SyncTarget, ref) Then
            If Not sync_confirm_info Is Nothing Then
                sync_confirm_info.ServicedItem = ref.Name
                sync_confirm_info.Info = "New Reference"
            Else
                RefAdd SyncTarget, ref
            End If
        End If
    Next ref
    
    '~~ Remove obsolete
    For Each ref In SyncTarget.VBProject.References
        If Not RefExists(SyncSource, ref) Then
            If Not sync_confirm_info Is Nothing Then
                sync_confirm_info.ServicedItem = ref.Name
                sync_confirm_info.Info = "Obsolete Reference"
            Else
                RefRemove SyncTarget, ref
            End If
        End If
    Next ref

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

Private Sub RefRemove( _
                ByRef rr_wb As Workbook, _
                ByVal rr_ref As Reference)
' -------------------------------------------------
' Removes Reference (rr_ref) from Workbook (rr_wb).
' -------------------------------------------------
    Dim ref As Reference
    
    With rr_wb.TargetWbWithSource
        For Each ref In .References
            If ref.Name = rr_ref.Name Then
                .References.Remove ref
                Exit Sub
            End If
        Next ref
    End With
    
End Sub

Private Sub RenameSheet(ByRef rs_wb As Workbook, _
                        ByVal rs_old_name As String, _
                        ByVal rs_new_name As String)
' ----------------------------------------------------
'
' ----------------------------------------------------
    Const PROC = "RenameSheet"
    
    On Error GoTo eh
    Dim sh  As Worksheet
    For Each sh In rs_wb.Worksheets
        If sh.Name = rs_old_name Then
            sh.Name = rs_new_name
            cLog.Entry = "Worksheet name changed to '" & rs_new_name & "'"
            Exit For
        End If
    Next sh

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub RenameWrkbkModule( _
                        ByRef rdm_wb As Workbook, _
                        ByVal rdm_new_name As String)
' ---------------------------------------------------
' Renames in Workbook (rdm_wb) the Workbook Module
' to (rdm_new_name).
' ---------------------------------------------------
    Const PROC = "RenameWrkbkModule"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    
    With rdm_wb.TargetWbWithSource
        For Each vbc In .VBComponents
            If vbc.Type = vbext_ct_Document Then
                If IsWrkbkComp(vbc) Then
                    cLog.ServicedItem = vbc.Name
                    vbc.Name = rdm_new_name
                    cLog.Entry = "Workbook Document Module renamed to '" & rdm_new_name & "'"
                    DoEvents
                    Exit For
                End If
            End If
        Next vbc
    End With
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub SheetsNew( _
       Optional ByRef sync_new_count As Variant = Nothing, _
       Optional ByRef sync_obsolete_count As Variant = Nothing, _
       Optional ByRef sync_confirm_info As clsLog)
' ---------------------------------------------------------------
' Synchronize new sheets in the source Workbook (SyncSource) with
' the target Workbook (SyncTarget).
' - When the optional new sheets counter (sync_new_count) is
'   provided, the new sheets are only counted
' - When the optional confirmation info (sync_confirm_info) is
'   provided only the syncronization infos are collect for being
'   confirmed.
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
    Const PROC = "SheetsNew"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim i       As Long
    Dim ws      As Worksheet
    Dim cSource As clsRaw
    Dim cTarget As clsComp
    
    For Each vbc In SyncSource.VBProject.VBComponents
        Set cSource = New clsRaw
        Set cSource.Wrkbk = SyncSource
        cSource.CompName = vbc.Name
        If CompExists(SyncTarget, vbc.Name) Then GoTo next_vbc
        If vbc.Type <> vbext_ct_Document Then GoTo next_vbc
        If Not IsSheetComp(vbc) Then GoTo next_vbc
        
        Set cTarget = New clsComp
        Set cTarget.Wrkbk = SyncTarget
        If cTarget.ExistsBySheetName(cSource.SheetName(vbc.Name)) Then GoTo next_vbc
            
        '~~ The sheet does not exist in the target Workbook neither under its Name
        '~~ nor under its CodeName.
        If VarType(sync_new_count) = vbLong Then
            '~~ This is just the first call for counting the potentially new sheets
            sync_new_count = sync_new_count + 1
        ElseIf Not sync_confirm_info Is Nothing Then
            '~~ This is just the second call for the collection of the sync confirmation info
            sync_confirm_info.ServicedItem = vbc.Name
            If lSheetsNew > 0 Or lSheetsObsolete > 0 Then
                If Not RestrictRenameAsserted Then
                    bAmbigous = True
                    sync_confirm_info.Info = "New Worksheet ambigous (sync denied unless sheet rename is asserted restricted)"
                Else
                    bAmbigous = False
                    sync_confirm_info.Info = "New Worksheet (will be copied from source Workbook)"
                End If
            Else
                sync_confirm_info.Info = "New Worksheet (will be copied from source Workbook)"
            End If
        Else
            '~~ This is the third call for getting the syncronizations done
            '~~ The new sheet is copied to the corresponding position in the target Workbook
            For i = 1 To cSource.Wrkbk.Worksheets.Count
                Set ws = cSource.Wrkbk.Worksheets(i)
                If ws.CodeName = cSource.CompName Then
                    If i = 1 _
                    Then ws.Copy Before:=SyncTarget.Sheets(1) _
                    Else ws.Copy After:=SyncTarget.Sheets(i - 1)
                    cLog.Entry = "New Workseet '" & cSource.CompName & "(" & ws.Name & ")' copied."
                    ' Exit For
                End If
            Next i
        End If
        Set cTarget = Nothing
        Set cSource = Nothing
next_vbc:
    Next vbc
       
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub SheetsObsolete( _
            Optional ByRef sync_obsolete_count As Variant = Nothing, _
            Optional ByRef sync_confirm_info As clsLog)
' --------------------------------------------------------------------
' Remove sheets in the target (SyncTarget) which are regarded
' obsolete because they do not exist in the target Workbook
' (SyncTarget) neither under their Name nor theit CodeName.
' - When the optional obsolet sheets counter (sync_obsolete_count)
'   is provided, the obsolete sheets are only counted
' - When the optional confirmation info (sync_confirm_info) is
'   provided only the syncronization infos are collect for being
'   confirmed.
' A Worksheet is finally only regarded a obsolete when
' A) it exists in the source Workbook neither under its CodeName nor
'    its Name a n d  it had been confirmed that name changes on sheets
'    are restricted to either or but never both at once.
' B) the number of new sheets is 0
' Note: This procedure is called three times
' 1. To count the sheets indicated obsolete
' 2. To get the removal of the obsolete sheets confirmed
' 3. To remove the obsolete sheets
' -----------------------------------------------------------------
    Const PROC = "SheetsObsolete"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim ws      As Worksheet
    Dim cSource As clsRaw
    Dim cTarget As clsComp
    
    For Each vbc In SyncTarget.VBProject.VBComponents
        If Not vbc.Type = vbext_ct_Document Then GoTo next_vbc
        If Not IsSheetComp(vbc) Then GoTo next_vbc
        
        Set cSource = New clsRaw
        Set cSource.Wrkbk = SyncSource
        cSource.CompName = vbc.Name
        If cSource.Exists Then GoTo next_vbc
        Set cTarget = New clsComp
        Set cTarget.Wrkbk = SyncTarget
        cTarget.CompName = vbc.Name
        If cSource.ExistsBySheetName(cTarget.SheetName(vbc.Name)) Then GoTo next_vbc
        
        '~~ Sheet Document Module exists in source Workbook
        '~~ neither under its Name nor under its CodeName
        cLog.ServicedItem = vbc.Name
        If VarType(sync_obsolete_count) = vbLong Then
            '~~ This is just the first call for counting the potentially new sheets
            sync_obsolete_count = sync_obsolete_count + 1
        ElseIf Not sync_confirm_info Is Nothing Then
            '~~ This is just the second call for the collection of the sync confirmation info
            sync_confirm_info.ServicedItem = vbc.Name
            If lSheetsNew > 0 Or lSheetsObsolete > 0 Then
                If Not RestrictRenameAsserted Then
                    bAmbigous = True
                    sync_confirm_info.Info = "Obsolete Worksheet ambigous (sync denied unless sheet rename is asserted restricted)"
                Else
                    bAmbigous = False
                    sync_confirm_info.Info = "Obsolete Worksheet (will be copied from source Workbook)"
                End If
            Else
                sync_confirm_info.Info = "Obsolete Worksheet (will be removed)"
            End If
        Else
            '~~ This is a Worksheet with no corresponding component and no corresponding sheet in the source Workbook.
            '~~ Because it has been asserted that sheets are never renamed by Name and CodeName at once
            '~~ this Worksheet is regarded obsolete for sure and will thus now be removed
            For Each ws In SyncTarget.Worksheets
                If ws.CodeName = vbc.Name Then
                    '~~ This is the obsolete sheet to be removed
                    ws.Delete
                    cLog.Entry = "Obsolete Worksheet deleted"
                    Exit For
                End If
            Next ws
        End If
        Set cTarget = Nothing
        Set cSource = Nothing
next_vbc:
    Next vbc
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub SheetsCodeNameChange( _
                  Optional ByRef sync_confirm_info As clsLog)
' ------------------------------------------------------------------
'
' ------------------------------------------------------------------
    Const PROC = "SheetsCodeNameChange"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsSource    As Worksheet
    Dim wsTarget    As Worksheet
    Dim vbc         As VBComponent
    
    For Each v In SourceSheets
        If Not TargetSheets.Exists(v) Then GoTo next_sheet
        
        Set wsSource = SourceSheets(v)
        Set wsTarget = TargetSheets(v)
        If wsTarget.CodeName = wsSource.CodeName Then GoTo next_sheet
        
        '~~ The sheet's CodName has changed while the sheet's Name remained unchanged
        If Not sync_confirm_info Is Nothing Then
            sync_confirm_info.ServicedItem = wsTarget.CodeName
            sync_confirm_info.Info = "Worksheet CodeName  " & wsTarget.CodeName & "  changes to  " & wsSource.CodeName
            sync_confirm_info.Info = "Worksheet code will be synced line by line"
        Else
            For Each vbc In SyncTarget.VBProject.VBComponents
                If vbc.Name = TargetSheets(v).CodeName Then
                    vbc.Name = wsSource.CodeName
                    '~~ When the sheet's CodeName has changed the sheet's code is synchronized line by line
                    '~~ because it is very likely code refers to the CodeName rather than to the sheet's Name or position
                    mSync.ByCodeLines sync_target_comp_name:=wsSource.CodeName _
                                    , sync_source_wb_full_name:=SyncSource.FullName
                    Exit For
                End If
            Next vbc
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

Private Sub SheetsNameChange( _
              Optional ByRef sync_confirm_info As clsLog)
' --------------------------------------------------------------
'
' --------------------------------------------------------------
    Const PROC = "SheetsameChange"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wsSource    As Worksheet
    Dim vbcSource   As VBComponent
    Dim vbcTarget   As VBComponent
    Dim ws          As Worksheet
    Dim sSheetName  As String
    Dim vbc         As VBComponent
    
    For Each v In SourceSheetComps
        If Not TargetSheetComps.Exists(v) Then GoTo next_comp
        Set vbcSource = SourceSheetComps(v)
        Set vbcTarget = TargetSheetComps(v)
        For Each ws In SyncTarget.Worksheets
            If ws.CodeName <> vbcSource.Name Then GoTo next_ws
            sSheetName = CompSheetName(SyncSource, vbcSource.Name)
            If ws.Name = sSheetName Then GoTo next_ws
            
            '~~ The sheet's Name has changed while the sheets CodeName remained unchanged
            If Not sync_confirm_info Is Nothing Then
                sync_confirm_info.ServicedItem = ws.CodeName
                sync_confirm_info.Info = "Worksheet Name  " & ws.Name & "  changes to  " & sSheetName
                sync_confirm_info.Info = "Worksheet code will be synced line by line"
            Else
                ws.Name = sSheetName
                '~~ When the sheet's name has changed the sheet's code is synchronized line by line
                '~~ for the case the name is used in the code
                For Each vbc In SyncTarget.VBProject.VBComponents
                    If vbc.Name = ws.CodeName Then
                        mSync.ByCodeLines sync_target_comp_name:=wsSource.CodeName _
                                        , sync_source_wb_full_name:=SyncSource.FullName
                        Exit For
                    End If
                Next vbc
            End If
next_ws:
        Next ws
next_comp:
    Next v

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub SheetsOrder( _
    Optional ByRef sync_confirm_info As clsLog = Nothing)
' -------------------------------------------------------
'
' -------------------------------------------------------
    Const PROC = "SheetsOrder"
    
    On Error GoTo eh
    Dim i           As Long
    Dim wsSource    As Worksheet
    Dim wsTarget    As Worksheet
    
    For i = 1 To SyncSource.Worksheets.Count
        Set wsSource = SyncSource.Worksheets(i)
        Set wsTarget = SyncTarget.Worksheets(i)
        If wsSource.Name <> wsTarget.Name Then
            '~~ Sheet position has changed
            If Not sync_confirm_info Is Nothing Then
                Stop ' pending confirmation info
            Else
                Stop ' pending implementation
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

Private Function CompSheetName( _
                         ByRef wb As Workbook, _
                         ByVal comp_name As String) As String
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.CodeName = comp_name Then
            CompSheetName = ws.Name
            Exit For
        End If
    Next ws
End Function

Private Sub SheetRemove(ByRef wb As Workbook, _
                        ByRef vbc As VBComponent)
' -----------------------------------------------
' Remove the sheet which corresponds with the
' VBComponent (vbc) by identinfying its Name.
' -----------------------------------------------
    Const PROC = "SheetRemove"
    
    On Error GoTo eh
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        If ws.CodeName = vbc.Name Then
            If ws.UsedRange.Columns.Count > 0 Or ws.UsedRange.Rows.Count > 0 Then
                If ws.Visible = xlSheetHidden Then
                    '~~ This Worksheet already contains data which might get lost when the sheet is deleted
                    '~~ For safety it will by copied as backup and hidden. However, this backup will be removed
                    '~~ with the next synchronization !!!
                    Stop
                    ws.Copy After:=wb.Worksheets.Count
                    ws.Name = ws.Name & "_Bkp"
                    ws.Visible = xlSheetHidden
                Else
                    '~~ This is a backup which now will be removed!
                    Stop
                    wb.Worksheets(ws.Name).Delete
                End If
            End If
            Exit For
        End If
    Next ws
    
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub TargetWbWithSource( _
                        ByRef sync_target_wb As Workbook, _
                        ByRef sync_source_wb As Workbook, _
               Optional ByVal restricted_sheet_rename_asserted As Boolean = False)
' --------------------------------------------------------------------------------
' Synchronizes a target Workbook (SyncTarget)
' with a source Workbook (SyncSource).
' --------------------------------------------------------------------------------
    Const PROC = "TargetWbWithSource"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim sConfirm        As String
    Dim v               As Variant
    Dim sMsg            As tMsg
    Dim sBttnCnfrmd     As String
    Dim sBttnTrmnt      As String
    Dim sBttnRestricted As String
    Dim cConf           As New clsLog
    Dim cllButtons      As Collection
    Dim sReply          As String
    Dim ts              As TextStream
    Dim ws              As Worksheet
    Dim vbc             As VBComponent
    
    RestrictRenameAsserted = restricted_sheet_rename_asserted
    Set SyncSource = sync_source_wb
    Set SyncTarget = sync_target_wb
        
    MaxLen lMaxLenComp, lMaxLenType
    
    Set dctChanged = New Dictionary
    CollectSyncObjects
    
    '~~ Count new and obsolete sheets
    lSheetsNew = 0
    lSheetsObsolete = 0
    SheetsNew sync_new_count:=lSheetsNew
    SheetsObsolete sync_obsolete_count:=lSheetsObsolete
    
    RestrictRenameAsserted = False
    bAmbigous = True
    bSyncDenied = True

    Do
        
        '~~ Collect all synchronization info and get them confirmed
        cConf.File = mFile.Temp
        If Not fso.FileExists(cConf.File) Then
            Set ts = fso.CreateTextFile(cConf.File)
            ts.Close
        End If
        cConf.CompMaxLen = CompMaxLen
        Set cConf.ServicedWrkbk = SyncSource
        References cConf
        
        CollectSyncConfInfo sync_new_count:=lSheetsNew _
                          , sync_obsolete_count:=lSheetsObsolete _
                          , sync_confirm_info:=cConf
        
        '~~ Get the collected info confirmed
        sConfirm = mFile.Txt(fso.GetFile(cConf.File))
        sMsg.Section(1).sText = sConfirm
        sMsg.Section(1).bMonspaced = True
        sMsg.Section(2).sText = "Please confirm the above synchronizations." & vbLf & _
                                "!! In case of any concerns terminate this syncronization service!"
        
        sBttnCnfrmd = "Synchronize" & vbLf & vbLf & fso.GetBaseName(SyncTarget.Name) & vbLf & " with " & vbLf & fso.GetBaseName(SyncSource.Name)
        sBttnTrmnt = "Terminate!" & vbLf & vbLf & "Synchronization denied" & vbLf & "because of concerns"
        sBttnRestricted = "Confirmed" & vbLf & "that Sheet rename" & vbLf & "is restricted" & vbLf & "(either Name  o r  CodeName)"
        
        If bAmbigous And Not RestrictRenameAsserted Then
            '~~ When Sheet names are regarded ambigous synchronization can only take place when it is confirmed
            '~~ that only either the CodeName or the Name is changed but not both. This ensures that sheets which cannot
            '~~ be mapped between the source and the target Workbook are either obsolete or new. The mapping inability
            '~~ may indicate that both sheet names (Name and CodeName) had been changed which cannot be synchronized
            '~~ because of the missing mapping.
            Set cllButtons = mMsg.Buttons(sBttnRestricted, sBttnTrmnt, vbLf)
            sMsg.Section(3).sText = "New and or Obsolete sheets are unclear when sheet's name change is " & _
                                    "not explicitely restricted to either the Name is change  o r  the CodeName " & _
                                    "but never both at once." & vbLf & _
                                    "Syncronization is denied unless this restriction is asserted!"
        Else
            Set cllButtons = mMsg.Buttons(sBttnCnfrmd, sBttnTrmnt, vbLf)
            sMsg.Section(3).sText = vbNullString
        End If
        For Each v In dctChanged
            cllButtons.Add v
        Next v
        
        sReply = mMsg.Dsply(msg_title:="Changes by synchronization require confirmation" _
                          , msg:=sMsg _
                          , msg_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sBttnTrmnt
                GoTo xt
            Case sBttnCnfrmd
                bSyncDenied = False
                Exit Do
            Case sBttnRestricted
                '~~ Collection of confirmation info is done again with this restriction now confirmed
                RestrictRenameAsserted = True
                If fso.FileExists(cConf.File) Then fso.DeleteFile (cConf.File)
            Case Else
                '~~ Display the requested changes
                Set cSource = dctChanged(sReply)
                cSource.DsplyAllChanges
        End Select
    Loop

    If Not bSyncDenied Then
        '~~ Synchronize ...
'        References
'        SheetsNew
'        SheetsObsolete
'        CompsNew
'        CompsObsolete
'        CompsChanged
    End If
    
xt: If fso.FileExists(cConf.File) Then fso.DeleteFile (cConf.File)
    Set cConf = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub CollectSyncConfInfo( _
                          ByVal sync_new_count As Long, _
                          ByVal sync_obsolete_count As Long, _
                          ByRef sync_confirm_info As clsLog)
' ------------------------------------------------------------
' Collect all confirmation information regarding sheet changes
' ------------------------------------------------------------
    SheetsCodeNameChange sync_confirm_info:=sync_confirm_info
    SheetsNameChange sync_confirm_info:=sync_confirm_info
    SheetsNew sync_confirm_info:=sync_confirm_info
    SheetsObsolete sync_confirm_info:=sync_confirm_info
    CompsNew sync_confirm_info:=sync_confirm_info
    CompsObsolete sync_confirm_info:=sync_confirm_info
    CompsChanged sync_confirm_info:=sync_confirm_info
End Sub

Private Sub CollectSyncObjects()

    Dim ws  As Worksheet
    Dim vbc As VBComponent
    Dim shp As Shape
    
    If SourceSheets Is Nothing Then Set SourceSheets = New Dictionary Else SourceSheets.RemoveAll
    If SourceSheetComps Is Nothing Then Set SourceSheetComps = New Dictionary Else SourceSheetComps.RemoveAll
    If TargetSheets Is Nothing Then Set TargetSheets = New Dictionary Else TargetSheets.RemoveAll
    If TargetSheetComps Is Nothing Then Set TargetSheetComps = New Dictionary Else TargetSheetComps.RemoveAll
    If SourceSheetShapes Is Nothing Then Set SourceSheetShapes = New Dictionary Else SourceSheetShapes.RemoveAll
    If TargetSheetShapes Is Nothing Then Set TargetSheetShapes = New Dictionary Else TargetSheetShapes.RemoveAll
    
    For Each ws In SyncSource.Worksheets
        If Not SourceSheets.Exists(ws.Name) Then SourceSheets.Add ws.Name, ws
        If Not SourceSheets.Exists(ws.CodeName) Then SourceSheets.Add ws.CodeName, ws
        For Each shp In ws.Shapes
            If Not SourceSheetShapes.Exists(shp.Name) Then SourceSheetShapes.Add shp.Name, ws
        Next shp
    Next ws
    For Each ws In SyncTarget.Worksheets
        If Not TargetSheets.Exists(ws.Name) Then TargetSheets.Add ws.Name, ws
        If Not TargetSheets.Exists(ws.CodeName) Then TargetSheets.Add ws.CodeName, ws
        For Each shp In ws.Shapes
            If Not TargetSheetShapes.Exists(shp.Name) Then TargetSheetShapes.Add shp.Name, ws
        Next shp
    Next ws
    For Each vbc In SyncSource.VBProject.VBComponents
        If IsSheetComp(vbc) Then
            If Not SourceSheetComps.Exists(vbc.Name) Then SourceSheetComps.Add vbc.Name, vbc
        End If
    Next vbc
    For Each vbc In SyncTarget.VBProject.VBComponents
        If IsSheetComp(vbc) Then
            If Not TargetSheetComps.Exists(vbc.Name) Then TargetSheetComps.Add vbc.Name, vbc
        End If
    Next vbc
    
End Sub

Private Function WbkGetOpen(ByVal go_wb_full_name As String) As Workbook
' ----------------------------------------------------------------------
' Returns an opened Workbook object named (go_wb_full_name) or Nothing
' when a file named (go_wb_full_name) not exists.
' ----------------------------------------------------------------------
    Const PROC = "WbkGetOpen"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If fso.FileExists(go_wb_full_name) Then
        If WbkIsOpen(wb_full_name:=go_wb_full_name) _
        Then Set WbkGetOpen = Application.Workbooks(go_wb_full_name) _
        Else Set WbkGetOpen = Application.Workbooks.Open(go_wb_full_name)
    End If
    
xt: Set fso = Nothing
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function WbkIsOpen( _
            Optional ByVal wb_base_name As String = vbNullString, _
            Optional ByVal wb_full_name As String = vbNullString) As Boolean
' -------------------------------------------------------------------------
' Returns TRUE when a Workbook either identified by its BaseName (wb_base_name)
' or by its full name (wb_full_name) is open. When the BaseName is
' provided in the current Excel instance, else in any Excel instance.
' -------------------------------------------------------------------------
    Const PROC = "WbkIsOpen"
    
    On Error GoTo eh
    Dim xlApp As Excel.Application
    
    If wb_base_name = vbNullString And wb_full_name = vbNullString Then GoTo xt
    
    With New FileSystemObject
        If wb_full_name <> vbNullString Then
            '~~ With the full name the open test spans all application instances
            If Not .FileExists(wb_full_name) Then GoTo xt
            If wb_base_name = vbNullString Then wb_base_name = .GetFileName(wb_full_name)
            On Error Resume Next
            Set xlApp = VBA.GetObject(wb_full_name).Application
            WbkIsOpen = Err.Number = 0
        Else
            On Error Resume Next
            wb_base_name = Application.Workbooks(wb_base_name).Name
            WbkIsOpen = Err.Number = 0
        End If
    End With

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

