Attribute VB_Name = "mFact"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mIs: Provides status results including relevant information.
' --------------------
'
' ----------------------------------------------------------------------------

Public Function IsServicedCommonComponent(ByVal i_comp As String) As Boolean
    IsServicedCommonComponent = CommonServiced.Components.Exists(i_comp)
End Function

Public Function IsServicedComponent(ByVal i_comp As String) As Boolean
    On Error Resume Next
    IsServicedComponent = Serviced.Wrkbk.VBProject.VBComponents(i_comp)
End Function

Public Function IsCommCompPublic(ByVal c_comp As String) As Boolean
    IsCommCompPublic = CommonPublic.Exists(c_comp)
End Function

Public Function CommCompHasBeenModifiedInTheServicedWorkbook(ByVal c_comp As clsComp, _
                                                    Optional ByRef c_mod_at_datetime_utc As String, _
                                                    Optional ByRef c_mod_in_wbk_fullname As String, _
                                                    Optional ByRef c_mod_on_machine As String) As Boolean
        
    With c_comp
        If .IsCommCompPublic Then
            If FSo.FileExists(.ExpFileFullName) Then
                Select Case True
                    Case c_mod_at_datetime_utc <> .ServicedLastModAt:   CommCompHasBeenModifiedInTheServicedWorkbook = False
                    Case c_mod_in_wbk_fullname <> .ServicedLastModIn:   CommCompHasBeenModifiedInTheServicedWorkbook = False
                    Case c_mod_on_machine <> .ServicedLastModOn:       CommCompHasBeenModifiedInTheServicedWorkbook = False
'                    Case .CodeExprtd.Meets(.CodePublic) = False:        CommCompHasBeenModifiedInTheServicedWorkbook = True
                    Case Else:                                          CommCompHasBeenModifiedInTheServicedWorkbook = True
                End Select
            End If
        End If
    End With

End Function


Private Sub Test_IsUsedCommonComp()
    Const PROC = "Test_IsUsedCommonComp"
    
    On Error GoTo eh
    Dim sLastModAtDatetime As String
    
    mCompMan.ServiceInitiate ThisWorkbook, PROC
    If CommonServiced.IsUsedCommComp("mBasic", sLastModAtDatetime) Then
        Debug.Print ErrSrc(PROC) & ": " & "Is a used Common Component"
        Debug.Print ErrSrc(PROC) & ": " & "Last modified at: " & sLastModAtDatetime
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_IsPendingReleaseCommComp()
    Const PROC = "Test_IsPendingReleaseCommComp"
    
    Dim sModAtDatetime      As String
    Dim sModExportFileName  As String
    Dim sModInWbkFullName   As String
    Dim sModOnMachine       As String
    
    If mCommComps.HasModificationPendingRelease("mBasic" _
                                              , sModAtDatetime _
                                              , sModExportFileName _
                                              , sModInWbkFullName _
                                              , sModOnMachine) Then
        Debug.Print ErrSrc(PROC) & ": " & "Is not pending release"
        Debug.Print ErrSrc(PROC) & ": " & "ModAtDatetime .........: " & sModAtDatetime
        Debug.Print ErrSrc(PROC) & ": " & "ModExportFileName .....: " & sModExportFileName
        Debug.Print ErrSrc(PROC) & ": " & "LastModInWrkbkFullName : " & sModInWbkFullName
        Debug.Print ErrSrc(PROC) & ": " & "ModOnMachine ..........: " & sModOnMachine
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

