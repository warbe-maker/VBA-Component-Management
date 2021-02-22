Attribute VB_Name = "mSync"
Option Explicit

Public Sub ByCodeLines(ByRef bcl_clone As clsComp, ByRef bcl_raw As clsRaw)
' --------------------------------------------------------------------------
' Synchronizes a Clone-VB-Project component's Worksheet code with its cor-
' responding raw component by replacing the clones code with the raw's code.
' Exception: When the Worksheet is new it requires information not available
'            through the Export-File. Thus the Raw Workbook is opened to
'            obtain the name of the Worksheet.
' --------------------------------------------------------------------------
    Const PROC = "ByCodeLines"

    On Error GoTo eh
    Dim i       As Long: i = 1
    Dim v       As Variant
    Dim dct     As Dictionary
    Dim ws      As Worksheet
    Dim wbRaw   As Workbook
    Dim sWsName As String
    
    With bcl_clone
        If Not mCompMan.CompExists(.Wrkbk, .CompName) Then
            Set wbRaw = mCompMan.WbkGetOpen(bcl_raw.HostFullName)
            For Each ws In wbRaw.Worksheets
                If ws.CodeName = .CompName Then
                    sWsName = ws.name
                    Exit For
                End If
            Next ws
            '~~ A not yet existing Worksheet must first be created
            Set ws = .Wrkbk.Worksheets.Add
            ws.name = sWsName
            Set ws = Nothing
        End If
    End With
    
    With bcl_clone.VBComp.CodeModule
        If .CountOfLines > 0 _
        Then .DeleteLines 1, .CountOfLines   ' Remove all lines from the cloned raw component
        
        Set dct = bcl_raw.CodeLines
        For Each v In dct    ' Insert the raw component's code lines
            .InsertLines i, dct(v)
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

Private Function CodeModuleIsEmpty(ByVal vbc As VBComponent) As Boolean
    With vbc.CodeModule
        CodeModuleIsEmpty = .CountOfLines = 0 Or .CountOfLines = 1 And Len(.Lines(1, 1)) < 2
    End With
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSync." & s
End Function

Private Function IsWrkbkComp(ByVal vbc As VBComponent) As Boolean
    
    Dim bSigned As Boolean
    On Error Resume Next
    bSigned = vbc.Properties("VBASigned").Value
    IsWrkbkComp = Err.Number = 0
    
End Function

Private Sub RemoveObsolete(ByRef ro_clone As clsComp)
' ----------------------------------------------------
'
' ----------------------------------------------------
    Const PROC = "RemoveObsolete"
    
    On Error GoTo eh
    Dim ws As Worksheet
    
    With ro_clone
        Select Case .VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_ClassModule
                '~~ The corresponding raw's Export File does not exist which indicates that the component no longer exists
                '~~ This is cruical in case the raw component just never had been exported! However, since the registration
                '~~ of the raw VB-Project is done by a service which makes sure all comopnents are exorted this should never
                '~~ be the case
                If Not CodeModuleIsEmpty(.VBComp) Then
                    .Wrkbk.VbProject.VBComponents.Remove .VBComp
                    cLog.Entry = "Obsolete component removed!"
                End If
            Case vbext_ct_Document
                If IsWrkbkComp(.VBComp) Then
                    cLog.Entry = "No Export File for the Workbook component is in fact impossible!"
                End If
                
                '~~ This is a Worksheet component with no corresponding Export File in the Raw-VB-Project
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

Private Sub CompUpdate( _
                 ByRef uc_clone As clsComp, _
                 ByRef uc_raw As clsRaw)
' ----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "CompUpdate"

    On Error GoTo eh
    With uc_clone
        '~~ The Export File of the Clone-VB-Project is not (no longer)
        '~~ identical with the Raw-VB-Project's corresponding Export File
        '~~ or the clone's corresponding Export-File does not exist which
        '~~ indicates that the raw component is new or has been renamed.
        Select Case .VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_ClassModule
                '~~ Standard Modules, Class Modules, and UserForms are updated by the import
                '~~ of the raw's Export-File
                mRenew.ByImport rn_wb:=.Wrkbk _
                              , rn_comp_name:=.CompName _
                              , rn_exp_file_full_name:=cRaw.ExpFileFullName
                cLog.Entry = "Renewed by import of " & cRaw.ExpFileFullName
                .VBComp.Export .ExpFileFullName
            Case vbext_ct_Document
                '~~ Data Modules cannot be updated by the import of the Export-File
                '~~ but by copying the code line by line
                If IsWrkbkComp(.VBComp) Then
                    cLog.Entry = "Programmatically synchronizing a code change in the Workbook component is not supported!"
                End If
                '~~ A code change in a Data Module can only be synchronized by replacing the component's code lines
                '~~ with those in the Raw-VB-Project
                mSync.ByCodeLines uc_clone, uc_raw
                cLog.Entry = "Change in raw component synchronized by re-writing the code line by line"
                .VBComp.Export .ExpFileFullName
            Case Else
                cLog.Entry = "Change of type '" & .TypeString & "' yet not supported!"
        End Select
    End With

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select

End Sub

Public Sub VbProject( _
               ByVal vb_clone_wb As Workbook, _
               ByVal vb_raw_wb As Workbook)
' --------------------------------------------------------
'
' --------------------------------------------------------
    Const PROC = "VbProject"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim cClone  As clsComp
    Dim cRaw    As clsRaw
    Dim vbc     As VBComponent
    Dim fl      As File
                                 
    For Each vbc In vb_raw_wb.VbProject.VBComponents
        cLog.ServicedItem = vbc.name
        If vbc.Type = vbext_ct_ActiveXDesigner Then
            cLog.Entry = "Type AxtiveXDesigner component is not supported"
            GoTo next_vbc
        End If
        
        If Not mCompMan.CompExists(vb_clone_wb, vbc.name) Then
            '~~ Add to the VB-Clone-Project the missing component
            If IsWrkbkComp(vbc) Then GoTo next_vbc
            CompAdd vb_clone_wb, vb_raw_wb, vbc.name
            GoTo next_vbc
        End If
        
        Set cClone = New clsComp
        Set cRaw = New clsRaw
        With cClone
            .Wrkbk = vb_clone_wb
            .CompName = vbc.name
            cLog.ServicedItem = .CompName
            .VBComp = vbc
            cRaw.HostFullName = vb_raw_wb.FullName
            cRaw.CompName = .CompName
            cRaw.ExpFileExtension = .ExpFileExtension
            cRaw.CloneExpFileFullName = .ExpFileFullName
            If Not fso.FileExists(.ExpFileFullName) Then
                .VBComp.Export .ExpFileFullName
            End If
            If cRaw.Changed Then
                '~~ Update the component of which the raw had changed
                CompUpdate cClone, cRaw
            Else
                cLog.Entry = "Already up-to-date!"
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

Private Function RefExists( _
                     ByVal re_wb As Workbook, _
                     ByVal re_name As String) As Boolean
' ------------------------------------------------------
'
' ------------------------------------------------------
    Dim ref As Reference
    
    For Each ref In re_wb.VbProject.References
        RefExists = ref.name = re_name
        If RefExists Then Exit Function
    Next ref

End Function

Private Sub RefRemove( _
                ByVal re_wb As Workbook, _
                ByVal re_name As String)
' ----------------------------------------
'
' ----------------------------------------
    Dim ref As Reference
    
    With re_wb.VbProject
        For Each ref In .References
            If ref.name = re_name Then
                .References.Remove ref
                Exit Sub
            End If
        Next ref
    End With
    
End Sub

Private Sub RefAdd(ByVal re_wb As Workbook, _
                   ByVal re_name As String)
' ------------------------------------------------
    re_wb.VbProject.References.AddFromFile re_name
End Sub

Private Sub CompRemoveObsolete( _
                    ByVal cr_clone_wb As Workbook, _
                    ByVal cr_raw_wb As Workbook)
' --------------------------------------------------
'
' --------------------------------------------------
End Sub

Private Sub CompExport(ByVal ce_wb As Workbook, _
              ByVal ce_comp_name As String)
    Dim cComp As New clsComp
    With cComp
        .Wrkbk = ce_wb
        .CompName = ce_comp_name
        .VBComp.Export .ExpFileFullName
    End With
End Sub
Private Sub CompAdd( _
              ByVal cr_clone_wb As Workbook, _
              ByVal cr_raw_wb As Workbook, _
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
        .Wrkbk = cr_raw_wb
        .CompName = cr_comp_name
        Set vbc = cr_raw_wb.VbProject.VBComponents(.CompName)
        Select Case vbc.Type
            Case vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_ClassModule
                cr_clone_wb.VbProject.VBComponents.Import .ExpFileFullName
                cLog.Entry = "Added by import of '" & "" & "'"
                CompExport cr_clone_wb, cr_comp_name
                End With
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

Private Sub NameAdd(ByVal na_target_wb As Workbook, _
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

Private Sub NameRemove()

End Sub

Private Sub NameUpdate()

End Sub
Private Sub CompSync( _
                    ByVal cr_clone_wb As Workbook, _
                    ByVal cr_raw_wb As Workbook, _
                    ByVal cr_comp_name As String)
' --------------------------------------------------
'
' --------------------------------------------------

End Sub

Private Function NameExists( _
                      ByVal ne_wb As Workbook, _
                      ByVal ne_nm As name) As Boolean
    Dim nm As name
    For Each nm In ne_wb.Names
        NameExists = nm.name = ne_nm.name
        If NameExists Then Exit For
    Next nm
End Function

