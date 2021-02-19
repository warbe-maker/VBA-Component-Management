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

Public Function IsWrkbkComp(ByVal vbc As VBComponent) As Boolean
    
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

Private Sub UpdateChanged( _
                    ByRef uc_clone As clsComp, _
                    ByRef uc_raw As clsRaw)
' ----------------------------------------------
'
' ----------------------------------------------
    Const PROC = "UpdateChanged"

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
               ByVal clone_project As Workbook, _
      Optional ByVal raw_project As String = vbNullString)
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
    
    '~~ Update all clone components where the raw had changed
    '~~ and remove all components where there is no Export File
    For Each vbc In clone_project.VbProject.VBComponents
        Set cClone = New clsComp
        Set cRaw = New clsRaw
        With cClone
            .Wrkbk = clone_project
            .CompName = vbc.name
            cLog.ServicedItem = .CompName
            .VBComp = vbc
            cRaw.HostFullName = raw_project
            cRaw.CompName = .CompName
            cRaw.ExpFileExtension = .ExpFileExtension
            cRaw.CloneExpFileFullName = .ExpFileFullName
                                        
            If .VBComp.Type = vbext_ct_ActiveXDesigner Then
                cLog.Entry = "This is a type '" & .TypeString & "' component which is not supported"
                GoTo next_vbc
            End If

            If Not fso.FileExists(cRaw.ExpFileFullName) Then
                RemoveObsolete cClone
                GoTo next_vbc
            End If
            
            If cRaw.Changed Then
                UpdateChanged cClone, cRaw
            Else
                cLog.Entry = "Already up-to-date!"
            End If
        End With
        
next_vbc:
    Set cClone = Nothing
    Set cRaw = Nothing
    Next vbc
    
    '~~ Adding componets yet not existing
    Set cRaw = New clsRaw
    cRaw.HostFullName = raw_project
    For Each fl In fso.GetFolder(cRaw.ExpFilePath).Files
        Debug.Print fl.Path
        Select Case fso.GetExtensionName(fl.Path)
            Case "bas", "cls", "frm"
                If Not mCompMan.CompExists(ce_wb:=clone_project, ce_comp_name:=fso.GetBaseName(fl.Path)) Then
                    '~~ Inporting a data module will fail!
                    cLog.ServicedItem = fso.GetBaseName(fl.Path)
                    On Error Resume Next
                    clone_project.VbProject.VBComponents.Import fl.Path
                    If Err.Number = 0 Then
                        cLog.Entry = "Component added by import of " & cRaw.ExpFileFullName
                    Else
                        cLog.Entry = "Component not added (adding data mudules not supported"
                    End If
                End If
        End Select
    Next fl
    
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

Private Sub CompAdd( _
                    ByVal cr_clone_wb As Workbook, _
                    ByVal cr_raw_wb As Workbook, _
                    ByVal cr_comp_name As String)
' --------------------------------------------------
'
' --------------------------------------------------
End Sub

Private Sub CompSync( _
                    ByVal cr_clone_wb As Workbook, _
                    ByVal cr_raw_wb As Workbook, _
                    ByVal cr_comp_name As String)
' --------------------------------------------------
'
' --------------------------------------------------

End Sub

