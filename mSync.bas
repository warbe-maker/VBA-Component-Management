Attribute VB_Name = "mSync"
Option Explicit

Private bSynchIssue     As Boolean
Private dctObsolete     As New Dictionary
Private dctNew          As New Dictionary
Private lMaxLenComp     As Long
Private lMaxLenType     As Long
Private dctSyncIssues   As Dictionary

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
    Dim nm      As Name
    
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
                        If InStr(nm.RefersTo, ws.Name & "!") <> 0 Then
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

Private Function CompExists( _
                      ByRef ce_wb As Workbook, _
                      ByVal ce_comp_name As String) As Boolean
' -----------------------------------------------------------
' Returns TRUE when the component (ce_comp_name) exists in
' the Workbook (ce_wb).
' -----------------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = ce_wb.VbProject.VBComponents(ce_comp_name).Name
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

Public Function Components(ByRef wb_clone As Workbook, _
                           ByRef wb_raw As Workbook) As Boolean
' -------------------------------------------------------------
' Synchronizes the productive (wb_clone) VB-Project with the
' development (wb_raw) VB-Project. When it returns False the
' synchronization is about to be terminated.
' -------------------------------------------------------------
    Const PROC = "Components"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim cClone      As clsComp
    Dim cRaw        As clsRaw
    Dim vbc         As VBComponent
    Dim sName       As String
    Dim cRawSheet   As clsSheet
    Dim cCloneSheet As clsSheet
    
    Dim sBttnAddAsNewSheet  As String: sBttnAddAsNewSheet = "Add the sheet as a new one" & vbLf & _
                                                            "Attention:" & vbLf & _
                                                            "One of the (productive!!!) sheets" & vbLf & _
                                                            "will be removed!"
    Dim sBttnTerminateSync  As String: sBttnTerminateSync = "Terminate this syncronization" & vbLf & _
                                                            "and make shure only the sheet's name" & vbLf & _
                                                            "o r  its CodeName is changed but not" & vbLf & _
                                                            "both at once!"
    For Each vbc In wb_raw.VbProject.VBComponents
        If vbc.Type = vbext_ct_ActiveXDesigner Then
            cLog.Entry = "Type AxtiveXDesigner component is not supported"
            GoTo next_vbc
        End If
        
        If Not CompExists(wb_clone, vbc.Name) _
        And Not vbc.Type = vbext_ct_Document Then
            '~~ Add a missing component to the VB-Clone-Project when it is
            '~~ neither a Workbook nor a Worksheet Document Module
            CompAdd wb_clone, wb_raw, vbc.Name
            cLog.ServicedItem = vbc.Name
            cLog.Entry = "Missing component added"
            GoTo next_vbc
        
        ElseIf Not CompExists(wb_clone, vbc.Name) _
        And IsWrkbkComp(vbc) Then
            '~~ When the Workbook Document Module does not exist it must be renamed
            RenameDocumentModule rdm_wb:=wb_clone _
                          , rdm_new_name:=vbc.Name _
                          , rdm_old_name:=sName
            cLog.ServicedItem = sName
            cLog.Entry = "Workbook component renamed to '" & vbc.Name & "'"
            GoTo next_vbc
        
        ElseIf IsSheetComp(vbc) Then
            Set cCloneSheet = New clsSheet
            Set cCloneSheet.Wrkbk = wb_clone
            Set cRawSheet = New clsSheet
            Set cRawSheet.Wrkbk = wb_raw
            
            If Not CompExists(wb_clone, vbc.Name) _
            And Not SheetExists(wb_clone, cRawSheet.Name(vbc.Name)) Then
                '~~ When Sheet Document Module neither exists under the Raw's CodeName
                '~~ nor under the Raw's Name both names may have been changed or the sheet is a new one.
                '~~ The latter is unlikely when the number of sheets had not changed. In case, this can
                '~~ only be clarified with the user
                If wb_clone.Worksheets.Count = wb_raw.Worksheets.Count Then
                    Select Case mMsg.Box(msg_title:="Sheet Document Module may be new or just renamed!" _
                                       , msg:="The Worksheet Document Module '" & vbc.Name & "' with sheet name '" & "" & "' " & _
                                              "neither exists in the productive (Clone-VB-Project) under this CodeName nor " & _
                                              "under its sheet name." & vbLf & vbLf & _
                                              "Please reply in accordance with the action to be taken." _
                                      , msg_buttons:=mMsg.Buttons(sBttnAddAsNewSheet, sBttnTerminateSync) _
                                       )
                        Case sBttnAddAsNewSheet
'                            CompAdd
                        Case Else:  GoTo xt
                    End Select
                                  
                ElseIf wb_clone.Worksheets.Count < wb_raw.Worksheets.Count Then
                    '~~ It's likely it is a new sheet!
                    wb_clone.Worksheets.Add cRawSheet.Name(vbc.Name)
                    Stop ' code name !!!!!!
                End If
            
            ElseIf Not CompExists(wb_clone, vbc.Name) _
            And SheetExists(wb_clone, cRawSheet.Name(vbc.Name)) Then
                '~~ A sheet which not exists under its CodeName
                '~~ but under its Name means that the CodeName had been changed
                RenameDocumentModule rdm_wb:=wb_clone _
                                   , rdm_new_name:=vbc.Name _
                                   , rdm_old_name:=sName
                cLog.ServicedItem = sName
                cLog.Entry = "Workbook component renamed to '" & vbc.Name & "'"
            ElseIf CompExists(wb_clone, vbc.Name) _
            And Not SheetExists(wb_clone, cRawSheet.Name(vbc.Name)) Then
                '~~ A sheet which exists under its CodeName
                '~~ but not under its Name means that its Name has changed
                wb_clone.Worksheets(cCloneSheet.Name(vbc.Name)).Name = cRawSheet.Name(vbc.Name)
            End If
            
            GoTo next_vbc
        End If

        
        '~~ Synchronice the code when the raw had changed
        Set cClone = New clsComp
        Set cRaw = New clsRaw
        With cClone
            Set .Wrkbk = wb_clone
            .CompName = vbc.Name
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
                    mSync.Controls wb_clone:=wb_clone, wb_raw:=wb_raw, ctrl_sheet_codename:=.CompName
                End If
            End If
        End With
        
next_vbc:
    Set cClone = Nothing
    Set cRaw = Nothing
    Next vbc
    Components = True
    
xt: Set fso = Nothing
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Sub CompRemove(ByRef cr_wb As Workbook, _
                       ByVal cr_comp As String)
' ----------------------------------------------------
'
' ----------------------------------------------------
    Const PROC = "CompRemove"
    
    On Error GoTo eh
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
                        ws.Name = ws.Name & "-bkp"
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

Private Function CompsNew( _
                    ByRef wb_clone As Workbook, _
                    ByRef wb_raw As Workbook, _
                    ByRef syncs As Dictionary)
' ---------------------------------------------------------
' Returns a Dictionary with all new non-Document-Modules in
' (wb_clone) Workbook, i.e. existing in (wb_raw) Workbook
' but not in (wb_clone).
' ---------------------------------------------------------
    Const PROC = "CompsNew"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim cComp       As clsComp
    Dim v           As Variant
    Dim lType       As vbcmType
    Dim vbc         As VBComponent
    Dim cCloneSheet As clsSheet
    Dim cRawSheet   As clsSheet
    Dim lNewSheets  As Long
    Dim lNewAdded   As Long
    
    '~~ Collect new Standard Modules, Class Modules, and UserForms
    Set cComp = New clsComp
    For Each v In cComp.CompType
        lType = v
        For Each vbc In wb_raw.VbProject.VBComponents
            If vbc.Type = lType Then
                If Not CompExists(wb_clone, vbc.Name) _
                And Not vbc.Type = vbext_ct_Document _
                Then mDct.DctAdd add_dct:=syncs _
                               , add_key:=vbc.Name & ":" & cComp.TypeString(vbc) & ":New" _
                               , add_item:=vbc _
                               , add_seq:=seq_ascending _
                               , add_order:=order_byitem
            End If
        Next vbc
    Next v

    '~~ Collect new Document Modules representing Worksheets
    Set dctSyncIssues = New Dictionary
    lNewSheets = wb_raw.Worksheets.Count - wb_clone.Worksheets.Count
    lNewAdded = 0
    For Each vbc In wb_raw.VbProject.VBComponents
        If Not CompExists(wb_clone, vbc.Name) _
        And vbc.Type = vbext_ct_Document _
        And Not IsWrkbkComp(vbc) Then
            '~~ A Worksheet is only regarded a new:
            '~~ 1. When the number of sheets in the Raw Workbook is greater than in the Clone Workbook.
            '~~ 2. When the sheet neither exists in the Clone Workbbook under its CodeName nor its Name.
            If wb_clone.Worksheets.Count < wb_raw.Worksheets.Count Then
                Set cCloneSheet = New clsSheet
                Set cCloneSheet.Wrkbk = wb_clone
                Set cRawSheet = New clsSheet
                Set cRawSheet.Wrkbk = wb_raw
                If Not cCloneSheet.ExistsByName(cRawSheet.Name(vbc.Name)) Then
                    If lNewAdded < lNewSheets Then
                        mDct.DctAdd add_dct:=syncs _
                                 , add_key:=vbc.Name & ":Worksheet:New" _
                                 , add_item:=vbc _
                                 , add_seq:=seq_ascending _
                               , add_order:=order_byitem
                        lNewAdded = lNewAdded + 1
                    Else
                        dctSyncIssues.Add vbc.Name, "There are sheets in '" & fso.GetBaseName(wb_raw.Name) & "' which appear new than the difference in number allows!"
                        dctSyncIssues.Add vbc.Name, "It cannot be assosiated with an existing Clone sheet when both, the Name and the CodeName are changed!"
                    End If
                End If
            End If
        End If
    Next vbc

xt: Set fso = Nothing
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
                         ByRef syncs As Dictionary)
' ---------------------------------------------------------------
' Returns a Dictionary with all non-Document-Modules obsolete
' in (wb_clone) Workbook, i.e. not existing in (wb_raw) Workbook.
' ---------------------------------------------------------------
    Const PROC = "CompsObsolete"
    
    On Error GoTo eh
    Dim cComp       As clsComp
    Dim v           As Variant
    Dim lType       As vbcmType
    Dim vbc         As VBComponent
    Dim cCloneSheet As clsSheet
    Dim cRawSheet   As clsSheet
    
    '~~ Collect obsolete Standard Modules, Class modules, and UserForms
    Set cComp = New clsComp
    For Each v In cComp.CompType
        lType = v
        For Each vbc In wb_clone.VbProject.VBComponents
            If vbc.Type = lType _
            And Not CompExists(wb_raw, vbc.Name) _
            And Not vbc.Type = vbext_ct_Document _
            Then mDct.DctAdd add_dct:=syncs _
                           , add_key:=vbc.Name & ":" & cComp.TypeString(vbc) & ":Obsolete" _
                           , add_item:=vbc _
                           , add_seq:=seq_ascending _
                           , add_order:=order_byitem
        Next vbc
    Next v

    '~~ Collect obsolete Standard Modules, Class modules, and UserForms
    For Each vbc In wb_clone.VbProject.VBComponents
        If vbc.Type = vbext_ct_Document _
        And IsSheetComp(vbc) _
        And Not CompExists(wb_raw, vbc.Name) Then
            '~~ Sheet Document Module not exists in Raw Workbook under its CodeName
            If wb_clone.Worksheets.Count > wb_raw.Worksheets.Count Then
                Set cCloneSheet = New clsSheet
                Set cCloneSheet.Wrkbk = wb_clone
                Set cRawSheet = New clsSheet
                Set cRawSheet.Wrkbk = wb_raw
                If Not cRawSheet.ExistsByName(cCloneSheet.Name(vbc.Name)) Then
                    mDct.DctAdd add_dct:=syncs _
                               , add_key:=vbc.Name & ":Worksheet:Obsolete" _
                               , add_item:=vbc _
                               , add_seq:=seq_ascending _
                               , add_order:=order_byitem
                End If
                Set cCloneSheet = Nothing
                Set cRawSheet = Nothing
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
    Dim cComp As clsComp
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

Private Sub RenameDocumentModule( _
                           ByRef rdm_wb As Workbook, _
                           ByVal rdm_old_name As String, _
                           ByVal rdm_new_name As String)
' -------------------------------------------------------
' Renames in Workbook (rdm_wb) the Data Module
' (rdm_old_name) to (rdm_new_name).
' -------------------------------------------------------
    Const PROC = "RenameDocumentModule"
    
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

Private Function AlignLeft(ByVal s As String, ByVal l As Long)
    AlignLeft = s & VBA.Space$(l - Len(s))
End Function

Public Sub VbProject(ByRef wb_clone As Workbook, _
                     ByRef wb_raw As Workbook)
' -------------------------------------------------
' Synchronizes the productive VB-Project (wb_clone)
' with the development VB-Project (wb_raw).
' -------------------------------------------------
    Const PROC = "VbProject"
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim sMsgConfirm As String
    Dim v           As Variant
    Dim sMsg        As tMsg
    Dim sBttnCnfrmd As String
    Dim sBttnTrmnt  As String
    Dim dctSyncs    As New Dictionary
    Dim a           As Variant
    
    MaxLen wb_clone, wb_raw, lMaxLenComp, lMaxLenType
  
    CompsNew wb_clone, wb_raw, dctSyncs
    CompsObsolete wb_clone, wb_raw, dctSyncs
        
    If dctNew.Count > 0 Or dctObsolete.Count > 0 Then
        For Each v In dctSyncs
            a = Split(v, ":")
            sMsgConfirm = AlignLeft(a(0), lMaxLenComp) & " " & AlignLeft(a(1), lMaxLenType) & " " & a(2)
        Next v
        sMsg.Section(1).sText = "Please confirm the following synchronization details." & vbLf & _
                                "In case of any concerns reply with No!"
        sMsg.Section(2).sText = sMsgConfirm
        sMsg.Section(2).bMonspaced = True
        
        sBttnCnfrmd = "Synchronize" & vbLf & vbLf & fso.GetBaseName(wb_clone.Name) & vbLf & " with " & vbLf & fso.GetBaseName(wb_raw.Name)
        sBttnTrmnt = "Do not synchronize!" & vbLf & vbLf & "(in case of any concerns)"
        
        
        If mMsg.Dsply(msg_title:="Please confirm the below synchronization details" _
                    , msg:=sMsg _
                    , msg_buttons:=mMsg.Buttons(sBttnCnfrmd, sBttnTrmnt) _
                     ) = sBttnTrmnt Then GoTo xt
    End If
    
    mSync.References wb_clone:=wb_clone, wb_raw:=wb_raw
    mSync.Components wb_clone:=wb_clone, wb_raw:=wb_raw
    
xt: Set dctSyncs = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Private Function CompNew(ByVal i As Long) As String
    Dim s As String
    CompNew = VBA.Space$(lMaxLenComp)
    If i <= dctNew.Count Then
        s = Split(dctNew.Keys()(i - 1), ":")(1)
        CompNew = s & VBA.Space$(lMaxLenComp - Len(s))
    End If
End Function

Private Function CompObsolete(ByVal i As Long) As String
    Dim s As String
    CompObsolete = VBA.Space$(lMaxLenComp)
    If i <= dctObsolete.Count Then
        s = Split(dctObsolete.Keys()(i - 1), ":")(1)
        CompObsolete = s & VBA.Space$(lMaxLenComp - Len(s))
    End If
End Function

Private Function TypeNew(ByVal i As Long) As String
    Dim s As String
    TypeNew = VBA.Space$(lMaxLenType)
    If i <= dctNew.Count Then
        s = Split(dctNew.Keys()(i - 1), ":")(0)
        TypeNew = s & VBA.Space$(lMaxLenType - Len(s))
    End If
End Function

Private Function TypeObsolete(ByVal i As Long) As String
    Dim s As String
    TypeObsolete = VBA.String$(lMaxLenType, " ")
    If i <= dctObsolete.Count Then
        s = Split(dctObsolete.Keys()(i - 1), ":")(0)
        TypeObsolete = s & VBA.Space$(lMaxLenType - Len(s))
    End If
End Function

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

