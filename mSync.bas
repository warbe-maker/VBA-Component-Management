Attribute VB_Name = "mSync"
Option Explicit

Public Function CloneAndRawProject( _
                    Optional ByRef cr_clone_wb As Workbook = Nothing, _
                    Optional ByVal cr_raw_name As String = vbNullString, _
                    Optional ByRef cr_raw_wb As Workbook = Nothing, _
                    Optional ByVal cr_confirm As Boolean = False) As Boolean
' --------------------------------------------------------------------------
' Returns True when the Clone-VB-Project and the Raw-VB-Project are valid.
' When bc_confirm is True a confirmation dialog is displayed. The dialog is
' also displayed when function is called with invalid arguments.
' --------------------------------------------------------------------------
    Const PROC                  As String = "CloneAndRawProject"
    Const CLONE_PROJECT         As String = "VB-Clone-Project"
    Const RAW_PROJECT           As String = "VB-Raw-Project"
    
    On Error GoTo eh
    Dim sBttCloneRawConfirmed   As String: sBttCloneRawConfirmed = "VB-Clone- and VB-Raw-Project" & vbLf & "Confirmed"
    Dim sBttnCloneProject       As String: sBttnCloneProject = "Select/change the" & vbLf & vbLf & CLONE_PROJECT & vbLf & " "
    Dim sBttnRawProject         As String: sBttnRawProject = "Configure/change the" & vbLf & vbLf & RAW_PROJECT & vbLf & " "
    Dim sBttnTerminate          As String
    sBttnTerminate = "Terminate providing a " & vbLf & _
                     "VB-Clone- and a VB-Raw-Project" & vbLf & _
                     "for being synchronized" & vbLf & _
                     "(sync service will be denied)"
    
    Dim fso             As New FileSystemObject
    Dim sMsg            As tMsg
    Dim sReply          As String
    Dim bWbClone        As Boolean
    Dim bWbRaw          As Boolean
    Dim sWbClone        As String
    Dim sWbRaw          As String ' either a full name or a registered raw project's basename
    Dim wbClone         As Workbook
    Dim wbRaw           As Workbook
    Dim cllButtons      As Collection
    Dim fl              As File
    
    If Not cr_clone_wb Is Nothing Then sWbClone = cr_clone_wb.FullName
    sWbRaw = cr_raw_name
    
    While (Not bWbClone Or Not bWbRaw) Or (cr_confirm And sReply <> sBttCloneRawConfirmed)
        If sWbClone = vbNullString Then
            sWbClone = "n o t  p r o v i d e d !"
        Else
            bWbClone = True
        End If
        
        If sWbRaw = vbNullString Then
            sWbRaw = "n o t  p r o v i d e d !"
        ElseIf Not fso.FileExists(sWbRaw) _
           And (Not mRawHosts.Exists(sWbRaw) _
                Or (mRawHosts.Exists(sWbRaw) And Not mRawHosts.IsRawVbProject(sWbRaw)) _
               ) Then
            sWbRaw = sWbRaw & ": i n v a l i d ! (neither a registered VB-Raw-Project nor a VB-Raw-Project's full name)"
        Else
            bWbRaw = True
        End If
    
        If bWbRaw And bWbClone And Not cr_confirm Then GoTo xt
        
        With sMsg
            .Section(1).sLabel = CLONE_PROJECT & ":"
            .Section(1).sText = sWbClone
            .Section(1).bMonspaced = True
            .Section(2).sLabel = RAW_PROJECT & ":"
            .Section(2).sText = sWbRaw
            .Section(2).bMonspaced = True
            
            If cr_confirm _
            Then .Section(3).sText = "Please confirm the provided VB-Clone- and VB-Raw-Project." _
            Else .Section(3).sText = "Please provide/complete the VB-Clone- (sync target) and the VB-Raw-Project (sync source)."
            
            .Section(3).sText = .Section(3).sText & vbLf & vbLf & _
                                "Attention!" & vbLf & _
                                "1. The '" & CLONE_PROJECT & "' must not be identical with the '" & RAW_PROJECT & "' and the two Workbooks must not have the same name." & vbLf & _
                                "2. Both VB-Projects/Workbook must exclusively reside in their parent Workbook" & vbLf & _
                                "3. Both Workbook folders must be subfolders of the configured '" & FOLDER_SERVICED & "'."

        End With
        
        '~~ Buttons preparation
        If Not bWbClone Or Not bWbRaw _
        Then Set cllButtons = mMsg.Buttons(sBttnRawProject, sBttnCloneProject, vbLf, sBttnTerminate) _
        Else Set cllButtons = mMsg.Buttons(sBttCloneRawConfirmed, vbLf, sBttnRawProject, sBttnCloneProject)
        
        sReply = mMsg.Dsply(msg_title:="Basic configuration of the Component Management (CompMan Addin)" _
                          , msg:=sMsg _
                          , msg_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sBttnCloneProject
                Do
                    If mFile.SelectFile(sel_filters:="*.xl*" _
                                      , sel_filter_name:="Workbook/VB-Project" _
                                      , sel_title:="Select the '" & CLONE_PROJECT & " to be synchronized with the '" & RAW_PROJECT & "'" _
                                      , sel_result:=fl _
                                       ) Then
                        sWbClone = fl.Path
                        Exit Do
                    End If
                Loop
                cr_confirm = True
                '~~ The change of the VB-Clone-Project may have made the VB-Raw-Project valid when formerly invalid
                sWbRaw = Split(sWbRaw, ": ")(0)
            Case sBttnRawProject
                Do
                    If mFile.SelectFile(sel_filters:="*.xl*" _
                                      , sel_filter_name:="Workbook/VB-Project" _
                                      , sel_title:="Select the '" & RAW_PROJECT & " as the synchronization source for the '" & CLONE_PROJECT & "'" _
                                      , sel_result:=fl _
                                       ) Then
                        sWbRaw = fl.Path
                        Exit Do
                    End If
                Loop
                cr_confirm = True
                '~~ The change of the VB-Raw-Project may have become valid when formerly invalid
                sWbClone = Split(sWbClone, ": ")(0)
            
            Case sBttCloneRawConfirmed: cr_confirm = False
            Case sBttnTerminate: GoTo xt
                
        End Select
        
    Wend ' Loop until the confirmed or configured basic configuration is correct
    
xt: If bWbClone Then
        If sWbClone <> cr_clone_wb.FullName Then Set cr_clone_wb = mCompMan.WbkGetOpen(sWbClone)
    End If
    If bWbRaw Then
        Set cr_raw_wb = mCompMan.WbkGetOpen(sWbRaw)
        cr_raw_name = fso.GetBaseName(cr_raw_wb.FullName)
    End If
    CloneAndRawProject = bWbClone And bWbRaw
    Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function



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
                                        
            If cRaw.Changed Then
                '~~ Update the component of which the raw had changed
                UpdateChanged cClone, cRaw
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

Private Sub CompAdd( _
                    ByVal cr_clone_wb As Workbook, _
                    ByVal cr_raw_wb As Workbook, _
                    ByVal cr_comp_name As String)
' --------------------------------------------------
'
' --------------------------------------------------
    Dim cSource As New clsComp
    Dim vbc     As VBComponent
    
    With cSource
        .Wrkbk = cr_raw_wb
        .CompName = cr_comp_name
        Set vbc = cr_raw_wb.VbProject.VBComponents(.CompName)
        Select Case vbc.Type
            Case vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_ClassModule
                cr_clone_wb.VbProject.VBComponents.Import .ExpFileFullName
                cLog.Entry = "Added by import of '" & "" & "'"
            
            Case vbext_ct_Document
                If Not IsWrkbkComp(vbc) Then
                    '~~ This new component is a Worksheet which is specifically delicate
                    '~~ 1. It cannot be simply copied from the raw to the clone because any referred names
                    '~~    would refer back to the raw source
                    '~~ 2. The code cannot be imported but has to be transferred from the Export-File line-by-line
                    
                End If
        End Select
    End With
    
    cLog.Entry = "Added as new Worksheet"
    cLog.Entry = "Code transferred from Export-File '" & "" & "'"
End Sub

Private Sub CompSync( _
                    ByVal cr_clone_wb As Workbook, _
                    ByVal cr_raw_wb As Workbook, _
                    ByVal cr_comp_name As String)
' --------------------------------------------------
'
' --------------------------------------------------

End Sub



