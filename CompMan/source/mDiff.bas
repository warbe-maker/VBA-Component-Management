Attribute VB_Name = "mDiff"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mDiff: Provides titles, buttons, difference checks and
' ====================== display for Common Component code differences,
' thereby unifying code difference detection, disply via WinMerge and
' corresponding dialogs.
'
' PendingVersusServicedCode        Returns TRUE when the code differs.
' PendingVersusServicedCodeBttn    Button caption
' PendingVersusServicedCodeDsply   Displays the difference between the current
'                                  pending release code and the serviced
'                                  component's CodeModule.
' PendingVersusServicedExport      Returns TRUE when the code differs.
' PendingVersusServicedExportBttn  Button caption
' PendingVersusServicedExportDsply Displays the difference between the current
'                                  pending release code and the serviced
'                                  component's Export-File
' PublicVersusPendingRelease       Returns TRUE when the code differs
' PublicVersusPendingReleaseBttn   Button caption
' PublicVersusPendingReleaseDsply  Displays the difference between the current
'                                  public code and the pending release code.
' PublicVersusServicedCode         Returns TRUE when the code differs.
' PublicVersusServicedCodeBttn     Button caption
' PublicVersusServicedCodeDsply    Displays the difference between the current
'                                  public code versus the serviced component's
'                                  CodeModule.
' ServicedCodeVersusPublic         Returns TRUE when the code differs
' ServicedCodeVersusPublicBttn     Button  caption
' ServicedCodeVersusPublicDsply    Displays the difference between the current
'                                  public code and the serviced component's
'                                  code in the CodeModule
' ServicedExportVersusPublic       Returns TRUE when the code differs
' ServicedExportVersusPublicBttn   Button  caption
' ServicedExportVersusPublicDsply  Displays the difference between the current
'                                  public code and the serviced component's
'                                  Export-File
' ----------------------------------------------------------------------------

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mDiff." & sProc
End Function

Public Function PendingVersusServicedCode(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the pending code differs from the serviced component's
' (p_comp) code in the CodeModule.
' ----------------------------------------------------------------------------
        
    With p_comp
        PendingVersusServicedCode = .CodePnding.DiffersFrom(.CodeExprtd)
    End With
    
End Function

Public Function PendingVersusServicedCodeBttn(ByVal p_comp As String) As String
    PendingVersusServicedCodeBttn = "Display difference" & vbLf & _
                                    mBasic.Spaced(p_comp) & vbLf & _
                                    "Pending versus current code" & vbLf & _
                                    "(VBComponent.CodeModule)"
End Function

Public Sub PendingVersusServicedCodeDsply(ByVal p_comp As clsComp)
' ----------------------------------------------------------------------------
' Displays the code difference between the pending release code and the
' current serviced component's (p_comp) Export-File.
' ----------------------------------------------------------------------------
    Const PROC = "PendingVersusServicedCodeDsply"
    
    On Error GoTo eh
    
    With p_comp
        .CodePnding.DsplyDiffs d_left_this_file_name:="CodePending" _
                             , d_left_this_file_title:=mDiff.SourcePending(.CompName) _
                             , d_rght_vrss_code:=.CodeCrrent _
                             , d_rght_vrss_file_name:="CurrentCodeModule" _
                             , d_rght_vrss_file_title:=mDiff.SourceServicedCode(.CompName)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function PendingVersusServicedExport(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the pending code differs from the serviced Export-File's
' code.
' ----------------------------------------------------------------------------
        
    With p_comp
        PendingVersusServicedExport = .CodePnding.DiffersFrom(.CodeExprtd)
    End With
    
End Function

Public Function PendingVersusServicedExportBttn(ByVal p_comp As String) As String
    PendingVersusServicedExportBttn = "Display difference" & vbLf & _
                                      mBasic.Spaced(p_comp) & vbLf & _
                                      "Pending versus current serviced" & vbLf & _
                                      "(Export-File)"
End Function

Public Sub PendingVersusServicedExportDsply(ByVal p_comp As clsComp)
' ----------------------------------------------------------------------------
' Displays the code difference between the pending release code and the
' current serviced component's (p_comp) Export-File.
' ----------------------------------------------------------------------------
    Const PROC = "PendingVersusServicedExportDsply"
    
    On Error GoTo eh
    
    With p_comp
        .CodePnding.DsplyDiffs d_left_this_file_name:="PendingRelease" _
                             , d_left_this_file_title:=mDiff.SourcePending(.CompName) _
                             , d_rght_vrss_code:=.CodeExprtd _
                             , d_rght_vrss_file_name:="LastModExport" _
                             , d_rght_vrss_file_title:=mDiff.SourceServicedExport(.CompName)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function PublicVersusPendingRelease(ByVal p_comp As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the public code differs from the pending code.
' ----------------------------------------------------------------------------
    
    With CommonPending
        Set .CodePublic = New clsCode
        With .CodePublic
            .CompName = p_comp
            .Source = CommonPublic.LastModExpFile(p_comp)
        End With
        Set .CodePnding = New clsCode
        With .CodePnding
            .CompName = p_comp
            .Source = CommonPending.LastModExpFile(p_comp)
        End With
        PublicVersusPendingRelease = .CodePublic.DiffersFrom(.CodePnding)
    End With
    
End Function

Public Function PublicVersusPendingReleaseBttn(ByVal p_comp As String) As String
    PublicVersusPendingReleaseBttn = "Display difference" & vbLf & _
                                     mBasic.Spaced(p_comp) & vbLf & _
                                     "Public versus Pending"
End Function

Public Sub PublicVersusPendingReleaseDsply(ByVal p_comp As String)
    Const PROC = "PublicVersusPendingReleaseDsply"
    
    On Error GoTo eh
    
    With CommonPending
        Set .CodePublic = New clsCode
        With .CodePublic
            .CompName = p_comp
            .Source = CommonPublic.LastModExpFile(p_comp)
        End With
        Set .CodePnding = New clsCode
        With .CodePnding
            .CompName = p_comp
            .Source = CommonPending.LastModExpFile(p_comp)
        End With
        .CodePublic.DsplyDiffs d_left_this_file_name:="PublicLastModExpFile" _
                             , d_left_this_file_title:=mDiff.SourcePublic(p_comp) _
                             , d_rght_vrss_code:=.CodePnding _
                             , d_rght_vrss_file_name:="CodePending" _
                             , d_rght_vrss_file_title:=mDiff.SourcePending(p_comp)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function PublicVersusServicedCodeBttn(ByVal p_comp As String) As String
    PublicVersusServicedCodeBttn = "Display difference" & vbLf & _
                                   mBasic.Spaced(p_comp) & vbLf & _
                                   "Public versus current" & vbLf & _
                                   "(CodeModule)"
End Function

Public Function PublicVersusServicedCode(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the serviced component's code (CodeModule) differs from
' the current public Common Component's code (Export-File).
' ----------------------------------------------------------------------------
    
    With p_comp
        PublicVersusServicedCode = .CodePublic.DiffersFrom(.CodeCrrent)
    End With
    
End Function

Public Sub PublicVersusServicedCodeDsply(ByVal p_comp As clsComp)
' ----------------------------------------------------------------------------
' Displays the code difference between the pending release code and the
' current serviced component's (p_comp) Export-File.
' ----------------------------------------------------------------------------
    Const PROC = "PublicVersusServicedCodeDsply"
    
    On Error GoTo eh
    
    With p_comp
        .CodePublic.DsplyDiffs d_left_this_file_name:="PublicExpFile" _
                             , d_left_this_file_title:=mDiff.SourcePublic(.CompName) _
                             , d_rght_vrss_code:=.CodeCrrent _
                             , d_rght_vrss_file_name:="ServicedCodeModule" _
                             , d_rght_vrss_file_title:=mDiff.SourceServicedCode(.CompName)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function PublicVersusServicedExport(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the pending code differs from the serviced Export-File's
' code.
' ----------------------------------------------------------------------------
    
    With p_comp
        PublicVersusServicedExport = .CodePnding.DiffersFrom(.CodeExprtd)
    End With
    
End Function

Public Function PublicVersusServicedExportBttn(ByVal p_comp As String) As String
    PublicVersusServicedExportBttn = "Display difference" & vbLf & _
                                      mBasic.Spaced(p_comp) & vbLf & _
                                      "Public versus current" & vbLf & _
                                      "(Export-File)"
End Function

Public Sub PublicVersusServicedExportDsply(ByVal p_comp As clsComp)
' ----------------------------------------------------------------------------
' Displays the code difference between the pending release code and the
' current serviced component's (p_comp) Export-File.
' ----------------------------------------------------------------------------
    Const PROC = "PublicVersusServicedExportDsply"
    
    On Error GoTo eh
    
    With p_comp
        .CodePublic.DsplyDiffs d_left_this_file_name:="PublicExpFile" _
                             , d_left_this_file_title:=mDiff.SourcePublic(.CompName) _
                             , d_rght_vrss_code:=.CodeExprtd _
                             , d_rght_vrss_file_name:="ServicedExport" _
                             , d_rght_vrss_file_title:=mDiff.SourceServicedExport(.CompName)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function ServicedExportVersusPublic(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the serviced component's (p_comp) Export-File differs from
' the public one in the Common-Components folder.
' ----------------------------------------------------------------------------
    
    Dim sComp   As String
    Dim arrDiff As Variant
    Dim v       As Variant
    
    With p_comp
        sComp = .CompName
        If .CodeExprtd.DiffersFrom(.CodePublic, arrDiff) Then
            ServicedExportVersusPublic = True
            With Servicing
                For Each v In arrDiff
                    .Log(sComp) = "Code difference " & v
                Next v
            End With
        End If
    End With
    
End Function

Public Function ServicedCodeVersusServicedExport(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the serviced component's (p_comp) code in the CodeModule
' differs from the public one in the Common-Components folder.
' ----------------------------------------------------------------------------
    Const PROC = ""
    
    On Error GoTo eh
    Dim sComp   As String
    Dim arrDiff As Variant
    Dim v       As Variant
    
    With p_comp
        sComp = .CompName
        If .CodeCrrent.DiffersFrom(.CodeExprtd, arrDiff) Then
            ServicedCodeVersusServicedExport = True
            With Servicing
                If ArryIsAllocated(arrDiff) Then
                    For Each v In arrDiff
                        .Log(sComp) = "Code difference " & v
                    Next v
                End If
            End With
        End If
    End With
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function ServicedCodeVersusPublic(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the serviced component's (p_comp) code in the CodeModule
' differs from the public one in the Common-Components folder.
' ----------------------------------------------------------------------------
    Const PROC = "ServicedCodeVersusPublic"
    
    On Error GoTo eh
    Dim sComp   As String
    Dim arrDiff As Variant
    Dim v       As Variant
    
    With p_comp
        sComp = .CompName
        If .CodeCrrent.DiffersFrom(.CodePublic, arrDiff) Then
            ServicedCodeVersusPublic = True
            If ArryIsAllocated(arrDiff) Then
                With Servicing
                    For Each v In arrDiff
                        .Log(sComp) = "Code difference " & v
                    Next v
                End With
            End If
        End If
    End With
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function ServicedExportVersusPublicBttn() As String
    ServicedExportVersusPublicBttn = "Display difference" & vbLf & _
                                     "Public versus current serviced" & vbLf & _
                                     "(Export-File)"
End Function

Public Sub ServicedExportVersusPublicDsply(ByVal p_comp As clsComp)
' ----------------------------------------------------------------------------
' Public versus serviced export
' ----------------------------------------------------------------------------
    Const PROC = "ServicedExportVersusPublicDsply"
    
    On Error GoTo eh
    
    With p_comp
        .CodeExprtd.DsplyDiffs d_left_this_file_name:="CommonComponentServiced" _
                             , d_left_this_file_title:=mDiff.SourceServicedExport(.CompName) _
                             , d_rght_vrss_code:=.CodePublic _
                             , d_rght_vrss_file_name:="CommonComponentPublic" _
                             , d_rght_vrss_file_title:=mDiff.SourcePublic(.CompName)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function ServicedExportVersusServicedCode(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the serviced component's (p_comp) Export-File differs from
' the public one in the Common-Components folder.
' ----------------------------------------------------------------------------
    Const PROC = "ServicedExportVersusServicedCode"
    
    On Error GoTo eh
    Dim sComp   As String
    Dim arrDiff As Variant
    Dim v       As Variant
    
    With p_comp
        sComp = .CompName
        If .CodeCrrent.DiffersFrom(.CodeExprtd, arrDiff) Then
            ServicedExportVersusServicedCode = True
            If .IsCommon Then
                With Servicing
                    .Log(sComp) = "Common Component modified: " & v
                    If ArryIsAllocated(arrDiff) Then
                        For Each v In arrDiff
                            .Log(sComp) = v
                        Next v
                    End If
                End With
            End If
        End If
    End With
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Source(ByVal s_str As String, _
                        ByVal s_comp As String) As String
    
    Source = s_str
    If s_comp <> vbNullString _
    Then Source = Source & " of  " & mBasic.Spaced(s_comp)
                                 
End Function

Public Function SourcePublic(ByVal p_comp As String) As String
    SourcePublic = Source("Public (Export-File)", p_comp)
End Function

Public Function SourcePending(ByVal p_comp As String) As String
    SourcePending = Source("Pending (Export-File)", p_comp)
End Function

Public Function SourceServicedCode(Optional ByVal s_comp As String = vbNullString) As String
    SourceServicedCode = Source("Serviced (CodeModule)", s_comp)
End Function

Public Function SourceServicedExport(Optional ByVal s_comp As String = vbNullString) As String
    SourceServicedExport = Source("Serviced (Export-File)", s_comp)
End Function

