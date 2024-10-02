Attribute VB_Name = "mFact"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mIs: Provides status results including relevant information.
' --------------------
'
' ----------------------------------------------------------------------------

Public Function IsServicedCommonComponent(ByVal i_comp As String) As Boolean
    IsServicedCommonComponent = CompManDat.Components.Exists(i_comp)
End Function

Public Function IsServicedComponent(ByVal i_comp As String) As Boolean
    On Error Resume Next
    IsServicedComponent = Serviced.Wrkbk.VBProject.VBComponents(i_comp)
End Function

Public Function HasBeenRenamedByUpdateService(ByVal i_comp As String) As Boolean
' ------------------------------------------------------------------------------
' Returns True when a component's name (comp_name indicates that it is had been
' renamed by CompMan to enable an update (rename/import) service.
' ------------------------------------------------------------------------------
    HasBeenRenamedByUpdateService = InStr(i_comp, RENAMED_BY_COMPMAN) <> 0
End Function

Public Function CommCompHasModificationPendingRelease(ByVal p_comp As String, _
                                             Optional ByRef p_last_mod_at_datetime_utc As String, _
                                             Optional ByRef p_last_mod_export_filename As String, _
                                             Optional ByRef p_last_mod_in_wbk_fullname As String, _
                                             Optional ByRef p_last_mod_in_wbk_name As String, _
                                             Optional ByRef p_last_mod_on_machine As String) As Boolean
    
    If CommonPending Is Nothing Then Set CommonPending = New clsCommonPending
    CommCompHasModificationPendingRelease = CommonPending.Exists(p_comp _
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

Public Function IsCommComp(ByVal c_comp As String) As Boolean
' ----------------------------------------------------------------------------
' Returns True when the component (c_comp) is either known as a public Common
' Component (in the Common-Components folder) or a hosted Common Component
' (which yet has ever been released to public).
' ----------------------------------------------------------------------------
    If CommonPublic.Exists(c_comp) _
    Then IsCommComp = True _
    Else If Serviced.Hosted.Exists(c_comp) _
    Then IsCommComp = True
    
End Function

Public Function IsCommCompPublic(ByVal c_comp As String) As Boolean
    IsCommCompPublic = CommonPublic.Exists(c_comp)
End Function

Public Function CommCompHasBeenModifiedinTheServicedWorkbook(ByVal c_comp As clsComp, _
                                                    Optional ByRef c_mod_at_datetime_utc As String, _
                                                    Optional ByRef c_mod_in_wbk_fullname As String, _
                                                    Optional ByRef c_mod_on_machine As String) As Boolean
        
    With c_comp
        If .IsCommCompPublic Then
            If FSo.FileExists(.ExpFileFullName) Then
                CommCompHasBeenModifiedinTheServicedWorkbook = .CodeExprtd.Meets(.CodePublic) = False
                c_mod_at_datetime_utc = .ServicedLastModAt
                c_mod_in_wbk_fullname = .ServicedLastModIn
            End If
        End If
    End With

End Function


Private Sub Test_IsUsedCommonComp()
    Const PROC = "Test_IsUsedCommonComp"
    
    On Error GoTo eh
    Dim sLastModAtDatetime As String
    
    mCompMan.ServiceInitiate ThisWorkbook, PROC
    If CommCompUsed("mBasic", sLastModAtDatetime) Then
        Debug.Print ErrSrc(PROC) & ": " & "Is a used Common Component"
        Debug.Print ErrSrc(PROC) & ": " & "Last modified at: " & sLastModAtDatetime
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function CommCompHasEffectiveManageGap(ByVal c_comp As clsComp, _
                                              ByRef c_last_mod As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE a Workbook is opened or saved and a public used Common
' Component's code indicates a management gap which is the case when an
' export of a Common Component regarded modified is exported and becomes
' pending release - although the exported Common Component is in fact just
' an outdated one which requires an update.
' This may be the case when a Workbook "comes back to being serviced" and a
' used Common Component has been modified and released i the meantime.
'
' Background: While "under Component Management" any code modification results
'             in a new export file when the Workbbook is saved. Therefore,
'             when there is no or none identical Export file there must be a
'             management gap. This gap will not matter when the used public
'             Common Component's code is still identical with the curtent
'             public code. When not and the Workbook is saved under an active
'             component management and this will/would manifest a code
'             modification although its a meanwhile "outdated code". This can
'             only be prevented by updating the code/component in order to
'             re-start with managed with an up-to-date code base or by
'             disabling the export.
' ----------------------------------------------------------------------------
    Const PROC = "CommCompHasEffectiveManageGap"
    
    Dim bCodeCurrentDiffersFromCodePublic       As Boolean
    Dim bExportFileExists                       As Boolean
    Dim bCommCompIsPendingByServicedWorkbook    As Boolean
    Dim bLastModDateTimeIsAvailable             As Boolean
    Dim bCodeCurrentDiffersFromCodePending      As Boolean
    Dim sComp                                   As String
    
    With c_comp
        sComp = .CompName
        If Not IsCommCompPublic(sComp) Then GoTo xt                 ' not a public Common Component
        
        bCodeCurrentDiffersFromCodePublic = .CodeExprtd.Meets(.CodePublic) = False
        bExportFileExists = FSo.FileExists(.ExpFileFullName)
        bLastModDateTimeIsAvailable = CompManDat.LastModAt(sComp) <> vbNullString ' possibly first time serviced
        bCommCompIsPendingByServicedWorkbook = .CommCompIsPendingByServicedWorkbook
        If bCommCompIsPendingByServicedWorkbook Then
            bCodeCurrentDiffersFromCodePending = .CodeExprtd.Meets(.CodePnding) = False
        End If
        
        Select Case True
            Case bCodeCurrentDiffersFromCodePublic = False: GoTo xt   ' used code is identical with public
            Case bExportFileExists = True _
             And bLastModDateTimeIsAvailable = True
                Select Case True
                    Case CompManDat.LastModAt(sComp) = CommonPublic.LastModAt(sComp)
                        CommCompHasEffectiveManageGap = True
                    Case CompManDat.LastModAt(sComp) < CommonPublic.LastModAt(sComp)
                        CommCompHasEffectiveManageGap = True
                    Case CompManDat.LastModAt(sComp) > CommonPublic.LastModAt(sComp)
                        CommCompHasEffectiveManageGap = True
                End Select
                
            Case bCommCompIsPendingByServicedWorkbook
                Debug.Print ErrSrc(PROC) & ": " & .CompName & ": Is pending release by this serviced Workbook"
                If bCodeCurrentDiffersFromCodePending = False Then GoTo xt  ' the pending release code is identical with the
                Debug.Print ErrSrc(PROC) & ": " & .CompName & ": Code pending differs from current"
            Case Else
                CommCompHasEffectiveManageGap = True
        End Select
        
    End With
    
xt:

End Function

Public Function CommCompIsPendingByServicedWorkbook(ByVal c_comp As clsComp)
    
    Dim sLastModInwbk As String
    
    With c_comp
        If .CommCompHasModificationPendingRelease(, , , sLastModInwbk) Then
            CommCompIsPendingByServicedWorkbook = sLastModInwbk = Serviced.Wrkbk.FullName _
                                                  And .CodeExprtd.Meets(.CodePnding)
        End If
    End With
    
End Function

Private Sub Test_IsPendingReleaseCommComp()
    Const PROC = "Test_IsPendingReleaseCommComp"
    
    Dim sModAtDatetime      As String
    Dim sModExportFileName  As String
    Dim sModInWbkFullName   As String
    Dim sModInWbkName       As String
    Dim sModOnMachine       As String
    Dim b                   As Boolean
    
    If CommCompHasModificationPendingRelease("mBasic" _
                                           , sModAtDatetime _
                                           , sModExportFileName _
                                           , sModInWbkFullName _
                                           , sModInWbkName _
                                           , sModOnMachine) Then
        Debug.Print ErrSrc(PROC) & ": " & "Is not pending release"
        Debug.Print ErrSrc(PROC) & ": " & "ModAtDatetime ...: " & sModAtDatetime
        Debug.Print ErrSrc(PROC) & ": " & "ModExportFileName: " & sModExportFileName
        Debug.Print ErrSrc(PROC) & ": " & "LastModInWrkbkFullName : " & sModInWbkFullName
        Debug.Print ErrSrc(PROC) & ": " & "ModInWbkame .....: " & sModInWbkName
        Debug.Print ErrSrc(PROC) & ": " & "ModOnMachine ....: " & sModOnMachine
    Else
        Debug.Print ErrSrc(PROC) & ": " & "Is not pending release"
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
                                              , i_last_mod_atdatetime_utc _
                                              , i_last_mod_expfile_fullname_origin _
                                              , i_last_mod_inwbk_fullname _
                                              , i_last_mod_inwbk_name _
                                              , i_last_mod_on_machine)
                        
End Function

