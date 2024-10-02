Attribute VB_Name = "mIs"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mIs: Provides status results including relevant information.
' --------------------
'
' ----------------------------------------------------------------------------

Public Function CommCompPendingRelease(ByVal p_comp As String, _
                              Optional ByRef p_last_mod_at_datetime_utc As String, _
                              Optional ByRef p_last_mod_export_filename As String, _
                              Optional ByRef p_last_mod_in_wbk_fullname As String, _
                              Optional ByRef p_last_mod_in_wbk_name As String, _
                              Optional ByRef p_last_mod_on_machine As String) As Boolean
    
    If CommonPending Is Nothing Then Set CommonPending = New clsCommonPending
    CommCompPendingRelease = CommonPending.Exists(p_comp _
                                                , p_last_mod_at_datetime_utc _
                                                , p_last_mod_export_filename _
                                                , p_last_mod_in_wbk_fullname _
                                                , p_last_mod_in_wbk_name _
                                                , p_last_mod_on_machine)
    
End Function

Public Function CommCompUsed(ByVal c_comp As String, _
                    Optional ByRef c_last_mod_at_datetime As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the component (c_comp) is a registered used Common Component
' ----------------------------------------------------------------------------
    Const PROC = "CommCompUsed"
    
    If Serviced Is Nothing _
    Then Err.Raise mBasic.AppErr(1), ErrSrc(PROC), "Then function cannot be used prior any service has been initiated!"
    If CompManDat Is Nothing _
    Then Err.Raise mBasic.AppErr(2), ErrSrc(PROC), "Then function cannot be used prior any service has been initiated!"
    
    CommCompUsed = CompManDat.IsUsedCommComp(c_comp, c_last_mod_at_datetime)
    
End Function

Public Function CommComp(ByVal c_comp As String) As Boolean
' ----------------------------------------------------------------------------
' Returns True when the component (c_comp) is either known as a public Common
' Component or a hosted Common Component which yet has ever been released to
' public.
' ----------------------------------------------------------------------------
    
    If CommonPublic.Exists(c_comp) _
    Then CommComp = True _
    Else If Serviced.Hosted.Exists(c_comp) _
    Then CommComp = True
    
End Function

Public Function CommCompPublic(ByVal c_comp As String) As Boolean
    CommCompPublic = CommonPublic.Exists(c_comp)
End Function

Public Function CommCompModified(ByVal c_comp As clsComp) As Boolean
        
    With c_comp
        If mIs.CommCompPublic(.CompName) _
        Then CommCompModified = .CodeExported.DiffersFrom(.CodePublic, True)
    End With

End Function


Private Sub Test_IsUsedCommonComp()
    Const PROC = "Test_IsUsedCommonComp"
    
    On Error GoTo eh
    Dim sLastModAtDatetime As String
    
    mCompMan.InitiateService ThisWorkbook, PROC
    If CommCompUsed("mBasic", sLastModAtDatetime) Then
        Debug.Print "Is a used Common Component"
        Debug.Print "Last modified at: " & sLastModAtDatetime
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_IsPendingReleaseCommComp()
    Dim sModAtDatetime      As String
    Dim sModExportFileName  As String
    Dim sModInWbkFullName   As String
    Dim sModInWbkName       As String
    Dim sModOnMachine       As String
    Dim b                   As Boolean
    
    If CommCompPendingRelease("mBasic" _
                            , sModAtDatetime _
                            , sModExportFileName _
                            , sModInWbkFullName _
                            , sModInWbkName _
                            , sModOnMachine) Then
        Debug.Print "Is not pending release"
        Debug.Print "ModAtDatetime ...: " & sModAtDatetime
        Debug.Print "ModExportFileName: " & sModExportFileName
        Debug.Print "ModInWbkFullName : " & sModInWbkFullName
        Debug.Print "ModInWbkame .....: " & sModInWbkName
        Debug.Print "ModOnMachine ....: " & sModOnMachine
    Else
        Debug.Print "Is not pending release"
    End If
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMe" & "." & sProc
End Function

Public Function ProcInComp(ByVal p_wbk As Workbook, _
                           ByVal p_comp As String, _
                           ByVal p_proc As String, _
                           ByRef p_vbcm As CodeModule) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a proedure (p_proc) exists in a component (p_comp) in a
' Workbook's (p_wbk) VBProject.
' ----------------------------------------------------------------------------
    
    Dim vbcm    As CodeModule
    Dim iLine   As Long
    Dim sProc   As String
    Dim lKind   As vbext_ProcKind
    
    On Error Resume Next
    Set vbcm = p_wbk.VBProject.VBComponents(p_comp).CodeModule
    If Err.Number <> 0 Then Exit Function
    
    With vbcm
        iLine = .CountOfDeclarationLines + 1
        Do Until iLine >= .CountOfLines
            sProc = .ProcOfLine(iLine, lKind)
            If sProc = p_proc Then
                ProcInComp = True
                Set p_vbcm = vbcm
                Exit Do
            End If
            iLine = .ProcStartLine(sProc, lKind) + .ProcCountLines(sProc, lKind) + 1
        Loop
    End With
    
End Function

Public Function PublicCommonComponent(ByVal i_comp As String, _
                             Optional ByRef i_export_file_extention As String, _
                             Optional ByRef i_last_mod_atdatetime_utc As String, _
                             Optional ByRef i_last_mod_expfile_fullname_origin As String, _
                             Optional ByRef i_last_mod_inwbk_fullname As String, _
                             Optional ByRef i_last_mod_inwbk_name As String, _
                             Optional ByRef i_last_mod_on_machine As String) As Boolean
' ----------------------------------------------------------------------------
' When the component (i_comp) exists in the CommComps.dat Private Profile File
' the function returns TRUE inxluding all relevant values.
' ----------------------------------------------------------------------------

    If CommonPublic Is Nothing Then Set CommonPublic = New clsCommonPublic
    PublicCommonComponent = CommonPublic.Exists(i_comp _
                                              , i_export_file_extention _
                                              , i_last_mod_atdatetime_utc _
                                              , i_last_mod_expfile_fullname_origin _
                                              , i_last_mod_inwbk_fullname _
                                              , i_last_mod_inwbk_name _
                                              , i_last_mod_on_machine)
                        
End Function

