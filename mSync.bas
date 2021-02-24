Attribute VB_Name = "mSync"
Option Explicit

Private bSynchIssue As Boolean

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
                sWsName = ws.name
                Exit For
            End If
        Next ws
        '~~ A not yet existing Worksheet must first be created
        Set ws = bcl_clone_wb.Worksheets.Add
        ws.name = sWsName
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

Private Function CodeModuleIsEmpty(ByRef vbc As VBComponent) As Boolean
    With vbc.CodeModule
        CodeModuleIsEmpty = .CountOfLines = 0 Or .CountOfLines = 1 And Len(.Lines(1, 1)) < 2
    End With
End Function

Private Sub CompAdd( _
              ByRef cr_clone_wb As Workbook, _
              ByRef cr_raw_wb As Workbook, _
              ByVal cr_comp_name As String)
' ---------------------------------------------
' Add the component (cr_comp_name) found in the
' VB-Raw-Project (cr_raw_wb) which yet doesn't
' exist in the VB-Clone_project (cr_clone_wb).
' ---------------------------------------------
    Const PROC = "CompAdd"
    
    On Error GoTo eh
    Dim cSource As New clsComp
    Dim vbc     As VBComponent
    Dim ws      As Worksheet
    Dim i       As Long
    Dim nm      As name
    Dim cClone  As New clsComp
    
    With cSource
        Set .Wrkbk = cr_raw_wb
        .CompName = cr_comp_name
        Set vbc = cr_raw_wb.VbProject.VBComponents(.CompName)
        Select Case vbc.Type
            Case vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_ClassModule
                cr_clone_wb.VbProject.VBComponents.Import .ExpFileFullName
                cLog.Entry = .TypeString & " added by import of '" & "" & "'"
                CompExport cr_clone_wb, cr_comp_name
    
            Case vbext_ct_Document
                If Not IsWrkbkComp(vbc) Then
                    '~~ 1. Copy the sheet to the corresponding position in the target Workbook
                    For i = 1 To .Wrkbk.Worksheets.Count
                        Set ws = .Wrkbk.Worksheets(i)
                        If ws.CodeName = .CompName Then
                            If i = 1 _
                            Then ws.Copy Before:=cr_clone_wb.Sheets(i) _
                            Else ws.Copy After:=cr_clone_wb.Sheets(i - 1)
                            cLog.Entry = "Copied"
                            Exit For
                        End If
                    Next i
                    '~~ 2. Transfer all corresponding names
                    For Each nm In .Wrkbk.Names
                        If InStr(nm.RefersTo, ws.name & "!") <> 0 Then
                            '~~ Name refers to copied sheet
                            If Not NameExists(.Wrkbk, nm) Then
                                NameAdd na_target_wb:=.Wrkbk, na_name:=nm
                            End If
                        End If
                    Next nm
 
                End If
        End Select
    End With
    
    cLog.Entry = "Code transferred from Export-File '" & "" & "'"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub CompDataModuleRename( _
                           ByRef wr_wb As Workbook, _
                           ByVal wr_old_name As String, _
                           ByVal wr_new_name As String)
' -------------------------------------------------------
' Renames in Workbook (wr_wb) the Data Module
' (wr_old_name) to (wr_new_name).
' -------------------------------------------------------
    Const PROC = "CompDataModuleRename"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim cClone As New clsComp
    
    With wr_wb.VbProject
        For Each vbc In .VBComponents
            If vbc.Type = vbext_ct_Document Then
                If vbc.name = wr_old_name Then
                    vbc.name = wr_new_name
                    DoEvents
                    Exit For
                End If
            End If
        Next vbc
    End With
    
    '~~ Export
    With cClone
        Set .Wrkbk = wr_wb
        .CompName = wr_new_name
        .VBComp.Export .ExpFileFullName
    End With

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Function CompExists( _
                      ByRef ce_wb As Workbook, _
                      ByVal ce_comp_name As String) As Boolean
' -----------------------------------------------------------
' Returns TRUE when the component (ce_comp_name) exists in
' the Workbook (ce_wb).
' -----------------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = ce_wb.VbProject.VBComponents(ce_comp_name).name
    CompExists = Err.Number = 0
End Function

Private Sub CompExport(ByRef ce_wb As Workbook, _
                       ByVal ce_comp_name As String)
    Dim cComp As New clsComp
    With cComp
        Set .Wrkbk = ce_wb
        .CompName = ce_comp_name
        .VBComp.Export .ExpFileFullName
    End With
End Sub

Public Sub Components(ByRef wb_clone As Workbook, _
                      ByRef wb_raw As Workbook)
' ----------------------------------------------------
' Synchronizes the productive VB-Project (wb_clone)
' with the development VB-Project (wb_raw).
' ----------------------------------------------------
    Const PROC = "Components"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim cClone  As clsComp
    Dim cRaw    As clsRaw
    Dim vbc     As VBComponent
    Dim fl      As File
    Dim sName   As String
    
    For Each vbc In wb_raw.VbProject.VBComponents
        If vbc.Type = vbext_ct_ActiveXDesigner Then
            cLog.Entry = "Type AxtiveXDesigner component is not supported"
            GoTo next_vbc
        End If
        
        '~~ Add a missing component to the VB-Clone-Project
        If Not CompExists(wb_clone, vbc.name) Then
            '~~ If the missing component is the Workbook component it just has got
            '~~ a different name in the VB-Raw-Project and thus has to be renamed.
            If IsWrkbkComp(vbc) Then
                CompDataModuleRename wr_wb:=wb_clone _
                              , wr_new_name:=vbc.name _
                              , wr_old_name:=sName
                cLog.ServicedItem = sName
                cLog.Entry = "Workbook component renamed to '" & vbc.name & "'"
            Else
                CompAdd wb_clone, wb_raw, vbc.name
                cLog.ServicedItem = vbc.name
                cLog.Entry = "Missing component added"
            End If
            GoTo next_vbc
        End If
        
        '~~ Synchronice the code when the raw had changed
        Set cClone = New clsComp
        Set cRaw = New clsRaw
        With cClone
            Set .Wrkbk = wb_clone
            .CompName = vbc.name
            cLog.ServicedItem = .CompName
            Set .VBComp = vbc
            cRaw.HostFullName = wb_raw.FullName
            cRaw.CompName = .CompName
            cRaw.ExpFileExtension = .ExpFileExtension
            cRaw.CloneExpFileFullName = .ExpFileFullName
            cRaw.TypeString = .TypeString
            If Not fso.FileExists(.ExpFileFullName) Then
                .VBComp.Export .ExpFileFullName
                cLog.Entry = .TypeString & " exported to '" & .ExpFileFullName & "'"
            End If
            If cRaw.Changed Then
                '~~ Update the component of which the raw had changed
                CompUpdate cClone, cRaw
            End If
            If .VBComp.Type = vbext_ct_Document Then
                If Not IsWrkbkComp(.VBComp) Then
                    '~~ Worksheet !
                    mSync.Names wb_clone:=wb_clone, wb_raw:=wb_raw, wb_sheet_name:=SheetName(wb_raw, .CompName)
                    mSync.Controls wb_clone:=wb_clone, wb_raw:=wb_raw, ctrl_sheet_codename:=.CompName
                End If
            End If
        End With
        
next_vbc:
    Set cClone = Nothing
    Set cRaw = Nothing
    Next vbc
        
xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Sub CompRemove(ByRef cr_wb As Workbook, _
                       ByVal cr_comp As String)
' ----------------------------------------------------
'
' ----------------------------------------------------
    Const PROC = "CompRemove"
    
    On Error GoTo eh
    Dim vbc As VBComponent
    Dim cComp   As New clsComp
    Dim ws      As Worksheet
    
    With cComp
        Set .Wrkbk = cr_wb
        .CompName = cr_comp
        Select Case .VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_ClassModule
                .Wrkbk.VbProject.VBComponents.Remove .VBComp
                cLog.Entry = "Obsolete " & .TypeString & " removed!"
            Case vbext_ct_Document
                If IsWrkbkComp(.VBComp) Then
                    cLog.Entry = "No Export-File for the Workbook component is in fact impossible!"
                End If
                
                '~~ This is a Worksheet component with no corresponding Export-File in the Raw-VB-Project
                '~~ Because the Worksheet may contain data and has just been renamed, for safety it is
                '~~ backed-up and hidden.
                Stop ' pending implementation !
                For Each ws In .Wrkbk.Worksheets
                    If ws.CodeName = .CompName Then
                        ws.name = ws.name & "-bkp"
                        ws.Visible = xlSheetHidden
                        Exit For
                    End If
                Next ws
                cLog.Entry = "Obsolete Workshhet renamed (as backup) and hidden"
        End Select
    End With
xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

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
            If Not CompExists(cr_raw_wb, vbc.name) _
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

Private Function IsWrkbkComp(ByRef vbc As VBComponent) As Boolean
    
    Dim bSigned As Boolean
    On Error Resume Next
    bSigned = vbc.Properties("VBASigned").Value
    IsWrkbkComp = Err.Number = 0
    
End Function

Private Sub NameAdd(ByRef na_target_wb As Workbook, _
                    ByVal na_name As name)

    With na_name
        na_target_wb.Names.Add name:=na_name.name _
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
                      ByVal ne_nm As name) As Boolean
    Dim nm As name
    For Each nm In ne_wb.Names
        NameExists = nm.name = ne_nm.name
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
    Dim nm          As name
    Dim sh          As Worksheet
    Dim shRaw       As Worksheet
    Dim shClone     As Worksheet
    Dim sSheetRaw   As String
    Dim sSheetClone As String
        
    sSheetRaw = shRaw.name
    sSheetClone = shClone.name
    
    If wb_sheet_name <> vbNullString Then
        For Each nm In wb_clone
            If Not NameExists(wb_raw, nm) And InStr(nm.name, wb_sheet_name & "!") <> 0 Then
                wb_clone.Names(nm.name).Delete
                cLog.Entry = "Obsolete range name '" & nm.name & "' removed"
            End If
        Next nm
    Else
        For Each nm In wb_clone
            If Not NameExists(wb_raw, nm) Then
                wb_clone.Names(nm.name).Delete
                cLog.Entry = "Obsolete range name '" & nm.name & "' removed"
            End If
        Next nm
    End If

xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
    End Select
End Sub

Private Sub NameUpdate()

End Sub

Private Function SheetGet(ByRef sg_wb As Workbook, _
                          ByVal sg_codename As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In sg_wb.Worksheets
        If sh.CodeName = sg_codename Then
            Set SheetGet = sh
            Exit For
        End If
    Next sh
End Function

Private Function SheetName( _
                     ByRef sn_wb As Workbook, _
                     ByVal sn_codename As String) As String
' -------------------------------------------------------------
' Returns the sheet name of a sheet identified by its CodeName.
' -------------------------------------------------------------
    Dim sh As Worksheet
    For Each sh In sn_wb.Worksheets
        If sh.CodeName = sn_codename Then
            SheetName = sh.name
            Exit For
        End If
    Next sh
End Function

Private Function SheetCodeName( _
                     ByRef scn_wb As Workbook, _
                     ByVal scn_name As String) As String
' ------------------------------------------------------
' Returns the sheet's CodeName identified by its Name.
' ------------------------------------------------------
    Dim sh As Worksheet
    For Each sh In scn_wb.Worksheets
        If sh.name = scn_name Then
            SheetCodeName = sh.CodeName
            Exit For
        End If
    Next sh
End Function


Private Sub RefAdd(ByRef ra_wb As Workbook, _
                   ByVal ra_ref As Reference)
' ----------------------------------------------------
    ra_wb.VbProject.References.AddFromFile ra_ref.name
End Sub

Private Sub References(ByRef wb_clone As Workbook, _
                       ByRef wb_raw As Workbook)
' --------------------------------------------------
' Synchronizes the References in vb_clone with those
' in wb_raw
' --------------------------------------------------

    Dim ref As Reference
    
    '~~ Add missing to clone Workbook
    For Each ref In wb_raw.VbProject.References
        If Not RefExists(wb_clone, ref) Then
            RefAdd wb_clone, ref
        End If
    Next ref
    
    '~~ Remove obsolete
    For Each ref In wb_clone.VbProject.References
        If Not RefExists(wb_raw, ref) Then
            RefRemove wb_clone, ref
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
        RefExists = ref.name = re_ref.name
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
            If ref.name = rr_ref.name Then
                .References.Remove ref
                Exit Sub
            End If
        Next ref
    End With
    
End Sub

Public Sub VbProject(ByRef wb_clone As Workbook, _
                     ByRef wb_raw As Workbook)
' ----------------------------------------------------
' Synchronizes the productive VB-Project (wb_clone)
' with the development VB-Project (wb_raw).
' ----------------------------------------------------
    Const PROC = "VbProject"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim cClone  As clsComp
    Dim cRaw    As clsRaw
    Dim vbc     As VBComponent
    Dim fl      As File
    Dim sName   As String
    
    mSync.References wb_clone:=wb_clone, wb_raw:=wb_raw
    mSync.Components wb_clone:=wb_clone, wb_raw:=wb_raw
    
xt: Set fso = Nothing
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
            wb_base_name = Application.Workbooks(wb_base_name).name
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

