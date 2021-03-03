Attribute VB_Name = "mSync"
Option Explicit

Private bSynchIssue     As Boolean
Private dctObsolete     As New Dictionary
Private dctNew          As New Dictionary
Private lMaxLenComp     As Long
Private lMaxLenType     As Long
Private dctChanged      As Dictionary   ' Confirm buttons and clsRaw items to display changed

Private Function AlignLeft(ByVal s As String, ByVal l As Long)
    AlignLeft = s & VBA.Space$(l - Len(s))
End Function

Public Sub ByCodeLines(ByRef bcl_clone_wb As Workbook, _
                       ByVal bcl_comp_name As String, _
                       ByVal bcl_raw_host_full_name As String, _
                       ByRef bcl_raw_codelines As Dictionary)
' -----------------------------------------------------------
' Synchronizes a Clone-VB-Project component's Worksheet code
' with its corresponding raw component by replacing the
' clones code with the raw's code.
' ------------------------------------------------------------
    Const PROC = "ByCodeLines"

    On Error GoTo eh
    Dim i       As Long: i = 1
    Dim v       As Variant
    Dim ws      As Worksheet
    Dim wbRaw   As Workbook
    Dim sWsName As String
    
    If Not CompExists(bcl_clone_wb, bcl_comp_name) Then
        Set wbRaw = WbkGetOpen(bcl_raw_host_full_name)
        For Each ws In wbRaw.Worksheets
            If ws.CodeName = bcl_comp_name Then
                sWsName = ws.Name
                Exit For
            End If
        Next ws
        '~~ A not yet existing Worksheet must first be created
        Set ws = bcl_clone_wb.Worksheets.Add
        ws.Name = sWsName
        Set ws = Nothing
    End If
    
    With bcl_clone_wb.VbProject.VBComponents(bcl_comp_name).CodeModule
        If .CountOfLines > 0 _
        Then .DeleteLines 1, .CountOfLines   ' Remove all lines from the cloned raw component
        
        For Each v In bcl_raw_codelines    ' Insert the raw component's code lines
            Debug.Print bcl_raw_codelines(v)
            .InsertLines i, bcl_raw_codelines(v)
            i = i + 1
        Next v
    End With
                
xt: Exit Sub

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

Private Function CodeModuleIsEmpty(ByRef vbc As VBComponent) As Boolean
    With vbc.CodeModule
        CodeModuleIsEmpty = .CountOfLines = 0 Or .CountOfLines = 1 And Len(.Lines(1, 1)) < 2
    End With
End Function

Private Function CompMaxLen( _
                      ByRef wb_clone As Workbook, _
                      ByRef wb_raw As Workbook) As Long
' --------------------------------------------------------
'
' --------------------------------------------------------
    Dim vbc As VBComponent
    Dim ref As Reference
    Dim l   As Long
    For Each vbc In wb_clone.VbProject.VBComponents:    l = mBasic.Max(l, Len(vbc.Name)):   Next vbc
    For Each vbc In wb_raw.VbProject.VBComponents:      l = mBasic.Max(l, Len(vbc.Name)):   Next vbc
    For Each ref In wb_clone.VbProject.References:      l = mBasic.Max(l, Len(ref.Name)):   Next ref
    For Each ref In wb_raw.VbProject.References:        l = mBasic.Max(l, Len(ref.Name)):   Next ref
    CompMaxLen = l
End Function

Private Sub CompRemoveObsolete( _
                    ByRef cr_clone_wb As Workbook, _
                    ByRef cr_raw_wb As Workbook)
' --------------------------------------------------
'
' --------------------------------------------------
    Const PROC = "CompRemoveObsolete"
    
    On Error GoTo eh
    Dim vbc As VBComponent
    Dim cll As New Collection
    Dim v   As Variant
    
    With cr_clone_wb.VbProject
        For Each vbc In .VBComponents
            If Not CompExists(cr_raw_wb, vbc.Name) _
            Then cll.Add vbc
        Next vbc
        For Each v In cll
            Set vbc = v
            .VBComponents.Remove vbc
        Next v
    End With
    
xt: Set cll = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub CompsChanged( _
                   ByRef wb_clone As Workbook, _
                   ByRef wb_raw As Workbook, _
          Optional ByRef confirm As clsLog = Nothing)
' -----------------------------------------------------
' When syncs is provided, collect all components which
' had changed. I.e. their raw Export-File differs from
' the Clone Export-File
' else perform the syncronization.
' -----------------------------------------------------
    Const PROC = "CompsNew"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim cComp       As clsComp
    Dim v           As Variant
    Dim lType       As vbcmType
    Dim vbc1        As VBComponent
    Dim vbc2        As VBComponent
    Dim cClone      As clsComp
    Dim cRaw        As clsRaw
    Dim lNewSheets  As Long
    Dim lNewAdded   As Long
    Dim sType       As String
    
    '~~ Collect changed Standard Modules, Class Modules, and UserForms
    For Each vbc1 In wb_raw.VbProject.VBComponents
        
        If IsSheetComp(vbc1) Then GoTo next_vbc1
        Set cRaw = New clsRaw
        Set cRaw.Wrkbk = wb_raw
        cRaw.CompName = vbc1.Name
        If Not cRaw.Exists(wb_clone) Then GoTo next_vbc1
        
        Set cClone = New clsComp
        Set cClone.Wrkbk = wb_clone
        cClone.CompName = vbc1.Name
        cRaw.CloneExpFileFullName = cClone.ExpFileFullName
        If Not cRaw.Changed Then GoTo next_vbc1
        
        If Not confirm Is Nothing Then
            confirm.ServicedItem = vbc1.Name
            confirm.Info = cClone.TypeString & " changed"
            If Not dctChanged.Exists("Display" & vbLf & vbc1.Name & vbLf & "changes") _
            Then dctChanged.Add "Display" & vbLf & vbc1.Name & vbLf & "changes", cRaw
        Else
            mRenew.ByImport rn_wb:=wb_clone _
                          , rn_comp_name:=vbc1.Name _
                          , rn_exp_file_full_name:=cRaw.ExpFileFullName
            cLog.Entry = "Renewed by import of '" & cRaw.ExpFileFullName & "'"
        End If
        
        Set cClone = Nothing
        Set cRaw = Nothing
next_vbc1:
    Next vbc1

    '~~ Collect changed Document Modules representing Worksheets
    For Each vbc2 In wb_raw.VbProject.VBComponents
        If Not vbc2.Type = vbext_ct_Document Then GoTo next_sheet
        If Not IsSheetComp(vbc2) Then GoTo next_sheet
        
        Set cRaw = New clsRaw
        Set cRaw.Wrkbk = wb_raw
        cRaw.CompName = vbc2.Name
        If Not cRaw.Exists(wb_clone) Then GoTo next_sheet
        
        Set cClone = New clsComp
        Set cClone.Wrkbk = wb_clone
        cClone.CompName = vbc2.Name
        cRaw.CloneExpFileFullName = cClone.ExpFileFullName
        If Not cRaw.Changed Then GoTo next_sheet
        
        If Not confirm Is Nothing Then
            confirm.ServicedItem = vbc2.Name
            confirm.Info = "Worksheet changed (about to be updated with code from Export-File line-by-line)"
            If Not dctChanged.Exists("Display" & vbLf & vbc2.Name & vbLf & "changes") _
            Then dctChanged.Add "Display" & vbLf & vbc2.Name & vbLf & "changes", cRaw
        Else
            cLog.ServicedItem = vbc2.Name
            mSync.ByCodeLines bcl_clone_wb:=wb_clone _
                            , bcl_comp_name:=vbc2.Name _
                            , bcl_raw_host_full_name:=cRaw.Wrkbk.FullName _
                            , bcl_raw_codelines:=cRaw.CodeLines
            cLog.Entry = "Worksheet code updated line-by-line from Export-File '" & cRaw.ExpFileFullName & "'"
        End If
        Set cRaw = Nothing
        Set cClone = Nothing
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
                    ByRef wb_clone As Workbook, _
                    ByRef wb_raw As Workbook, _
           Optional ByRef confirm As clsLog = Nothing)
' ------------------------------------------------------
' Adds to Dictionary (syncs) all new Components, i.e.
' those not existing in Workbook (wb_raw).
' ------------------------------------------------------
    Const PROC = "CompsNew"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim v           As Variant
    Dim lType       As vbcmType
    Dim vbc1        As VBComponent
    Dim vbc2        As VBComponent
    Dim cComp       As clsComp
    Dim cClone      As clsComp
    Dim cRaw        As clsRaw
    Dim lNewSheets  As Long ' number of new sheets in the Raw Workbook
    Dim lNew        As Long ' sheets collected / added
    Dim i           As Long
    Dim ws          As Worksheet
    
    '~~ Collect/synchronize new Standard Modules, Class Modules, and UserForms
    Set cComp = New clsComp
    For Each vbc1 In wb_raw.VbProject.VBComponents
        If vbc1.Type = vbext_ct_Document _
        Or vbc1.Type = vbext_ct_ActiveXDesigner Then GoTo next_vbc1
        
        Set cRaw = New clsRaw
        Set cRaw.Wrkbk = wb_raw
        cRaw.CompName = vbc1.Name
        If CompExists(wb_clone, vbc1.Name) Then GoTo next_vbc1
        
        If Not confirm Is Nothing Then
            confirm.ServicedItem = vbc1.Name
            confirm.Info = "New " & cRaw.TypeString
        Else
            cLog.ServicedItem = vbc1.Name
        End If
        
        Set cRaw = Nothing
next_vbc1:
    Next vbc1

    '~~ Collect/synchronize new Document Modules representing Worksheets
    lNewSheets = wb_raw.Worksheets.Count - wb_clone.Worksheets.Count
    lNew = 0
    For Each vbc2 In wb_raw.VbProject.VBComponents
        Set cRaw = New clsRaw
        Set cRaw.Wrkbk = wb_raw
        cRaw.CompName = vbc2.Name
        If Not CompExists(wb_clone, vbc2.Name) _
        And vbc2.Type = vbext_ct_Document Then
            If IsWrkbkComp(vbc2) Then
                '~~ A Document Module representing the Workbook can't be new but obviously has bben renamed
                If Not confirm Is Nothing Then
                    confirm.ServicedItem = vbc2.Name
                    confirm.Info = "Worksheet to be renamed to '" & cRaw.CompName & "'"
                Else
                    RenameWrkbkModule wb_clone, cRaw.CompName
                End If
            ElseIf IsSheetComp(vbc2) Then
                '~~ A Worksheet is only regarded a new:
                '~~ 1. When the number of sheets in the Raw Workbook is greater than in the Clone Workbook.
                '~~ 2. When the sheet neither exists in the Clone Workbbook under its CodeName nor its Name.
                If wb_clone.Worksheets.Count < wb_raw.Worksheets.Count Then
                    Set cClone = New clsComp
                    Set cClone.Wrkbk = wb_clone
                    If Not cClone.ExistsBySheetName(cRaw.SheetName(vbc2.Name)) Then
                        If lNew < lNewSheets Then
                            If Not confirm Is Nothing Then
                                confirm.ServicedItem = vbc2.Name
                                confirm.Info = "New Worksheet"
                                lNew = lNew + 1
                            Else
                                '~~ 1. Copy the sheet to the corresponding position in the target Workbook
                                For i = 1 To cRaw.Wrkbk.Worksheets.Count
                                    Set ws = cRaw.Wrkbk.Worksheets(i)
                                    If ws.CodeName = cRaw.CompName Then
                                        If i = 1 _
                                        Then ws.Copy Before:=wb_clone.SheetsNameAndCodeName(1) _
                                        Else ws.Copy After:=wb_clone.SheetsNameAndCodeName(i - 1)
                                        cLog.Entry = "New Workseet '" & cRaw.CompName & "(" & ws.Name & ")' copied."
                                        ' Exit For
                                    End If
                                Next i
                            End If
                        Else
                            If Not confirm Is Nothing Then
                                confirm.ServicedItem = vbc2.Name
                                confirm.Info = "New Worksheet denied! Sheet's CodeName and Sheet's Name confused!"
                            End If
                            lNew = lNew + 1
                        End If
                    End If
                    Set cClone = Nothing
                End If
            End If
        End If
        Set cRaw = Nothing
next_vbc2:
    Next vbc2

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
                         ByRef wb_clone As Workbook, _
                         ByRef wb_raw As Workbook, _
                Optional ByRef confirm As clsLog = Nothing)
' -----------------------------------------------------------
' When (syncs) is provided obsolete components are collected
' for being confirmed, else obsolete components are removed -
' obsolete Worksheets renamed bkp and hidden.
' -----------------------------------------------------------
    Const PROC = "CompsObsolete"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim lType       As vbcmType
    Dim vbc         As VBComponent
    Dim cClone      As clsComp
    Dim cRaw        As clsRaw
    Dim sComp       As String
    Dim sType       As String
    Dim sBkpName    As String
    Dim ws          As Worksheet
    Dim lObsoleteSheets As Long
    Dim lSheetsRemoved  As Long
    
    '~~ Collect obsolete Standard Modules, Class modules, and UserForms
    Set cComp = New clsComp
    For Each v In cComp.CompType
        lType = v
        For Each vbc In wb_clone.VbProject.VBComponents
            Set cClone = New clsComp
            Set cClone.Wrkbk = wb_clone
            cClone.CompName = vbc.Name
            If vbc.Type = lType _
            And Not cClone.Exists(wb_raw) _
            And Not vbc.Type = vbext_ct_Document Then
                If Not confirm Is Nothing Then
                    confirm.ServicedItem = vbc.Name
                    confirm.Info = "Obsolete " & cClone.TypeString & " (will be removed)"
                Else
                    cLog.ServicedItem = vbc.Name
                    sType = cClone.TypeString
                    wb_clone.VbProject.VBComponents.Remove vbc
                    cLog.Entry = "Obsolete " & sType & " removed"
                End If
            End If
            Set cClone = Nothing
        Next vbc
    Next v

    '~~ Collect obsolete Worksheets
    lObsoleteSheets = wb_clone.Worksheets.Count - wb_raw.Worksheets.Count
    For Each vbc In wb_clone.VbProject.VBComponents
        Set cClone = New clsComp
        Set cClone.Wrkbk = wb_clone
        cClone.CompName = vbc.Name
        Set cRaw = New clsRaw
        Set cRaw.Wrkbk = wb_raw
        cRaw.CompName = vbc.Name
        If vbc.Type = vbext_ct_Document _
        And cClone.IsSheet _
        And Not cRaw.Exists Then
            '~~ Sheet Document Module not exists in Raw Workbook under its CodeName
            If lSheetsRemoved < lObsoleteSheets Then
                If Not cRaw.ExistsBySheetName(cClone.SheetName(vbc.Name)) Then
                    If Not confirm Is Nothing Then
                        confirm.ServicedItem = vbc.Name
                        confirm.Info = "Obsolete Worksheet (will be removed!)"
                        lSheetsRemoved = lSheetsRemoved + 1
                    Else
                        '~~ This is a Worksheet component with no corresponding Export-File in the Raw-VB-Project
                        '~~ Because the Worksheet may contain data and has just been renamed, for safety it is
                        '~~ backed-up and hidden.
                        Stop ' pending implementation !
                        For Each ws In wb_clone.Worksheets
                            If ws.CodeName = cRaw.CompName Then
                                sBkpName = ws.Name & "-bkp"
                                ws.Name = sBkpName
                                ws.Visible = xlSheetHidden
                                cLog.ServicedItem = cClone.CompName
                                cLog.Entry = "Worksheet renamed to '" & sBkpName & "' (for safety in case it contained data)"
                                Exit For
                            End If
                        Next ws
                    End If
                End If
                Set cClone = Nothing
                Set cRaw = Nothing
            Else
                '~~ More sheets to be removed the obsolete in number
                If Not confirm Is Nothing Then
                    confirm.ServicedItem = vbc.Name
                    confirm.Info = "Obsolete Worksheet will  n o t  be removed (CodeName and Name likely confused!)"
                    lSheetsRemoved = lSheetsRemoved + 1
                End If
            End If
        End If
    Next vbc

xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Sub CompUpdate( _
                 ByRef uc_clone As clsComp, _
                 ByRef uc_raw As clsRaw)
' ----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "CompUpdate"

    On Error GoTo eh
    With uc_clone
        '~~ The Export-File of the Clone-VB-Project is not (no longer)
        '~~ identical with the Raw-VB-Project's corresponding Export-File
        '~~ or the clone's corresponding Export-File does not exist which
        '~~ indicates that the raw component is new or has been renamed.
        Select Case .VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_ClassModule
                '~~ Standard Modules, Class Modules, and UserForms are updated by the import
                '~~ of the raw's Export-File
                mRenew.ByImport rn_wb:=.Wrkbk _
                              , rn_comp_name:=.CompName _
                              , rn_exp_file_full_name:=cRaw.ExpFileFullName
                cLog.Entry = "Renewed by import of '" & cRaw.ExpFileFullName & "'"
                .VBComp.Export .ExpFileFullName
                cLog.Entry = "Exported to '" & .ExpFileFullName & "'"
            Case vbext_ct_Document
                '~~ A code change in a Data Module can only be synchronized by replacing the component's code lines
                '~~ with those in the Raw-VB-Project
                mSync.ByCodeLines bcl_clone_wb:=.Wrkbk _
                                , bcl_comp_name:=.CompName _
                                , bcl_raw_host_full_name:=uc_raw.HostFullName _
                                , bcl_raw_codelines:=uc_raw.CodeLines
                cLog.Entry = "Change in raw component synchronized by re-writing " & uc_clone.VBComp.CodeModule.CountOfLines & " code lines"
                .VBComp.Export .ExpFileFullName
                cLog.Entry = "Exported to '" & .ExpFileFullName & "'"
            Case Else
                cLog.Entry = "Change of type '" & .TypeString & "' yet not supported!"
        End Select
    End With ' uc_clone

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select

End Sub

Private Sub Controls(ByRef wb_clone As Workbook, _
                     ByRef wb_raw As Workbook, _
            Optional ByVal ctrl_sheet_name As String = vbNullString, _
            Optional ByVal ctrl_sheet_codename As String = vbNullString)
' ----------------------------------------------------------------------
'
' ----------------------------------------------------------------------

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
                      ByRef wb_clone As Workbook, _
                      ByRef wb_raw As Workbook, _
                      ByRef max_len_comp As Long, _
                      ByRef max_len_type As Long) As Long
    Set cComp = New clsComp
    Set cComp.Wrkbk = wb_clone
    max_len_comp = cComp.MaxLenComp
    Set cComp = New clsComp
    Set cComp.Wrkbk = wb_raw
    max_len_comp = mBasic.Max(max_len_comp, cComp.MaxLenComp)
    max_len_type = cComp.MaxLenType
End Function

Private Sub NameAdd(ByRef na_target_wb As Workbook, _
                    ByVal na_name As Name)

    With na_name
        na_target_wb.Names.Add Name:=na_name.Name _
                             , RefersTo:=na_name.RefersTo _
                             , Visible:=.Visible _
                             , MacroType:=.MacroType _
                             , ShortcutKey:=.ShortcutKey _
                             , Category:=.Category _
                             , NameLocal:=.NameLocal _
                             , RefersToLocal:=.RefersToLocal
    End With
End Sub

Private Function NameExists( _
                      ByRef ne_wb As Workbook, _
                      ByVal ne_nm As Name) As Boolean
    Dim nm As Name
    For Each nm In ne_wb.Names
        NameExists = nm.Name = ne_nm.Name
        If NameExists Then Exit For
    Next nm
End Function

Public Sub Names(ByRef wb_clone As Workbook, _
                 ByRef wb_raw As Workbook, _
        Optional ByVal wb_sheet_name As String = vbNullString, _
        Optional ByVal wb_sheet_codename As String = vbNullString)
' ----------------------------------------------------------------
' Synchronize the names in Worksheet (wb_clone) with those in
' Workbook wb_raw) - when either a sheet's name (wb_sheet_name) or
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
        For Each nm In wb_clone
            If Not NameExists(wb_raw, nm) And InStr(nm.Name, wb_sheet_name & "!") <> 0 Then
                wb_clone.Names(nm.Name).Delete
                cLog.Entry = "Obsolete range name '" & nm.Name & "' removed"
            End If
        Next nm
    Else
        For Each nm In wb_clone
            If Not NameExists(wb_raw, nm) Then
                wb_clone.Names(nm.Name).Delete
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
    ra_wb.VbProject.References.AddFromFile ra_ref.Name
End Sub

Private Sub References(ByRef wb_clone As Workbook, _
                       ByRef wb_raw As Workbook, _
              Optional ByRef confirm As clsLog = Nothing)
' ---------------------------------------------------------
' Collect to be synchronized References (when confirm Not
' is Nothing) else synchronizes the References.
' ---------------------------------------------------------

    Dim ref As Reference
    
    '~~ Add missing to clone Workbook
    For Each ref In wb_raw.VbProject.References
        If Not RefExists(wb_clone, ref) Then
            If Not confirm Is Nothing Then
                confirm.ServicedItem = ref.Name
                confirm.Info = "New Reference"
            Else
                RefAdd wb_clone, ref
            End If
        End If
    Next ref
    
    '~~ Remove obsolete
    For Each ref In wb_clone.VbProject.References
        If Not RefExists(wb_raw, ref) Then
            If Not confirm Is Nothing Then
                confirm.ServicedItem = ref.Name
                confirm.Info = "Obsolete Reference"
            Else
                RefRemove wb_clone, ref
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
    
    For Each ref In re_wb.VbProject.References
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
    
    With rr_wb.VbProject
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

Private Sub RenameSheetModule( _
                        ByRef rdm_wb As Workbook, _
                        ByVal rdm_old_name As String, _
                        ByVal rdm_new_name As String)
' -----------------------------------------------------
' Renames in Workbook (rdm_wb) the Workbook Module to
' (rdm_new_name).
' ---------------------------------------------------
    Const PROC = "RenameWrkbkModule"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim cClone As New clsComp
    
    With rdm_wb.VbProject
        For Each vbc In .VBComponents
            If vbc.Type = vbext_ct_Document Then
                If vbc.Name = rdm_old_name Then
                    vbc.Name = rdm_new_name
                    DoEvents
                    Exit For
                End If
            End If
        Next vbc
    End With
    
    '~~ Export
    With cClone
        Set .Wrkbk = rdm_wb
        .CompName = rdm_new_name
        .VBComp.Export .ExpFileFullName
    End With

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
    Dim cClone As New clsComp
    
    With rdm_wb.VbProject
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

Private Sub SheetsNameAndCodeName( _
                            ByRef wb_clone As Workbook, _
                            ByRef wb_raw As Workbook, _
                   Optional ByRef confirm As clsLog = Nothing)
' ------------------------------------------------------------
' Collect to be synchronized Worksheets' Name and/or CodeName
' (when confirm Not is Nothing) else synchronizes the names.
' ------------------------------------------------------------
    Const PROC = "SheetsNameAndCodeName"
    
    On Error GoTo eh
    Dim dctRawComp      As New Dictionary
    Dim dctCloneComp    As New Dictionary
    Dim dctRawSheet     As New Dictionary
    Dim dctCloneSheet   As New Dictionary
    Dim ws              As Worksheet
    Dim vbc             As VBComponent
    Dim v               As Variant
    Dim wsRaw           As Worksheet
    Dim wsClone         As Worksheet
    Dim vbcRaw          As VBComponent
    Dim vbcClone        As VBComponent
    Dim sSheetName      As String
    
    For Each ws In wb_raw.Worksheets
        If Not dctRawSheet.Exists(ws.Name) Then dctRawSheet.Add ws.Name, ws
        If Not dctRawSheet.Exists(ws.CodeName) Then dctRawSheet.Add ws.CodeName, ws
    Next ws
    For Each ws In wb_clone.Worksheets
        If Not dctCloneSheet.Exists(ws.Name) Then dctCloneSheet.Add ws.Name, ws
        If Not dctCloneSheet.Exists(ws.CodeName) Then dctCloneSheet.Add ws.CodeName, ws
    Next ws
    For Each vbc In wb_raw.VbProject.VBComponents
        If IsSheetComp(vbc) Then
            If Not dctRawComp.Exists(vbc.Name) Then dctRawComp.Add vbc.Name, vbc
        End If
    Next vbc
    For Each vbc In wb_clone.VbProject.VBComponents
        If IsSheetComp(vbc) Then
            If Not dctCloneComp.Exists(vbc.Name) Then dctCloneComp.Add vbc.Name, vbc
        End If
    Next vbc
        
    '~~ Change CodeNames
    For Each v In dctRawSheet
        If Not dctCloneSheet.Exists(v) Then GoTo next_sheet
            Set wsRaw = dctRawSheet(v)
            Set wsClone = dctCloneSheet(v)
            If wsClone.CodeName <> wsRaw.CodeName Then
            '~~ The sheet's CodName has changed while the sheet's Name remained unchanged
            If Not confirm Is Nothing Then
                confirm.ServicedItem = wsClone.CodeName
                confirm.Info = "Worksheet CodeName  " & wsClone.CodeName & "  changes to  " & wsRaw.CodeName
            Else
                For Each vbc In wb_clone.VbProject.VBComponents
                    If vbc.Name = dctCloneSheet(v).CodeName Then
                        vbc.Name = wsRaw.CodeName
                        Exit For
                    End If
                Next vbc
            End If
        End If
next_sheet:
    Next v
    
    '~~ Change Sheet-Names
    For Each v In dctRawComp
        If Not dctCloneComp.Exists(v) Then GoTo next_comp
        Set vbcRaw = dctRawComp(v)
        Set vbcClone = dctCloneComp(v)
        For Each ws In wb_clone.Worksheets
            If ws.CodeName = vbcRaw.Name Then
                sSheetName = CompSheetName(wb_raw, vbcRaw.Name)
                If ws.Name <> sSheetName Then
                    '~~ The sheet's Name has changed while the sheets CodeName remained unchanged
                    If Not confirm Is Nothing Then
                        confirm.ServicedItem = wsClone.CodeName
                        confirm.Info = "Worksheet Name  " & ws.Name & "  changes to  " & sSheetName
                    Else
                        ws.Name = sSheetName
                    End If
                    Exit For
                End If
            End If
        Next ws
next_comp:
    Next v


xt: Set dctRawComp = Nothing
    Set dctCloneComp = Nothing
    Set dctRawSheet = Nothing
    Set dctCloneSheet = Nothing
    Exit Sub
    
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

Private Function SheetExists( _
                       ByRef se_wb As Workbook, _
                       ByVal se_name As String) As Boolean
' ----------------------------------------------------------------
' Returns True when the sheet (se_name) exists in Workbook (se_wb)
' ----------------------------------------------------------------
    Dim ws As Worksheet
    For Each ws In se_wb.Worksheets
        If ws.Name = se_name Then
            SheetExists = True: Exit For
        End If
    Next ws
End Function

Private Function SheetGet(ByRef sg_wb As Workbook, _
                          ByVal sg_codename As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In sg_wb.Worksheets
        If sh.CodeName = sg_codename Then
            Set SheetGet = sh: Exit For
        End If
    Next sh
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

Public Sub VbProject(ByRef wb_clone As Workbook, _
                     ByRef wb_raw As Workbook)
' -------------------------------------------------
' Synchronizes the productive VB-Project (wb_clone)
' with the development VB-Project (wb_raw).
' -------------------------------------------------
    Const PROC = "VbProject"
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim sConfirm    As String
    Dim v           As Variant
    Dim sMsg        As tMsg
    Dim sBttnCnfrmd As String
    Dim sBttnTrmnt  As String
    Dim a           As Variant
    Dim cConf       As New clsLog
    Dim cllButtons  As Collection
    Dim cRaw        As clsRaw
    Dim sReply      As String
    Dim ts          As TextStream
    
    MaxLen wb_clone, wb_raw, lMaxLenComp, lMaxLenType
    cConf.File = mFile.Temp
    If Not fso.FileExists(cConf.File) Then
        Set ts = fso.CreateTextFile(cConf.File)
        ts.Close
    End If
    cConf.CompMaxLen = CompMaxLen(wb_clone, wb_raw)
    Set cConf.ServicedWrkbk = wb_raw
    Set dctChanged = New Dictionary
    
    '~~ Collect all synchronization info and get them confirmed
'    References wb_clone, wb_raw, cConf
    SheetsNameAndCodeName wb_clone, wb_raw, cConf
'    CompsNew wb_clone, wb_raw, cConf
'    CompsObsolete wb_clone, wb_raw, cConf
'    CompsChanged wb_clone, wb_raw, cConf
        
    sConfirm = mFile.Txt(fso.GetFile(cConf.File))
    If sConfirm <> vbNullString Then
        sMsg.Section(1).sText = sConfirm
        sMsg.Section(1).bMonspaced = True
        sMsg.Section(2).sText = "Please confirm the above synchronizations." & vbLf & _
                                "!! In case of any concerns terminate this syncronization service!"
        
        sBttnCnfrmd = "Synchronize" & vbLf & vbLf & fso.GetBaseName(wb_clone.Name) & vbLf & " with " & vbLf & fso.GetBaseName(wb_raw.Name)
        sBttnTrmnt = "Terminate!" & vbLf & vbLf & "Synchronization denied" & vbLf & "because of concerns"
            
        Set cllButtons = mMsg.Buttons(sBttnCnfrmd, sBttnTrmnt, vbLf)
        For Each v In dctChanged
            cllButtons.Add v
        Next v
        Do
            sReply = mMsg.Dsply(msg_title:="Changes by synchronization require confirmation" _
                              , msg:=sMsg _
                              , msg_buttons:=cllButtons _
                               )
            Select Case sReply
                Case sBttnTrmnt: GoTo xt
                Case sBttnCnfrmd: Exit Do
                Case Else
                    '~~ Display the requested changes
                    Set cRaw = dctChanged(sReply)
                    cRaw.DsplyAllChanges
            End Select
        Loop
    
'        References wb_clone, wb_raw
        SheetsNameAndCodeName wb_clone, wb_raw
'        CompsNew wb_clone, wb_raw
'        CompsObsolete wb_clone, wb_raw
'        CompsChanged wb_clone, wb_raw
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

