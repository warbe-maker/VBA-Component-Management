Attribute VB_Name = "mDiff"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mDiff: Provides titles, buttons, difference checks and
' ====================== display for Common Component code differences,
' thereby unifying WinMerge titles and button captions.
'
' PendingVersusServicedCode        Returns TRUE when the code differs.
' PendingVersusServicedCodeBttn    Button caption
' PendingVersusServicedExport      Returns TRUE when the code differs.
' PendingVersusServicedExportBttn  Button caption
' PendingVersusServicedExportDsply Displays the difference between the current
'                                  pending release code and the serviced
'                                  component's Export-File
' PublicVersusPendingRelease       Retunrs TRUE when the code differs
' PublicVersusPendingReleaseBttn   Button caption
' PublicVersusPendingReleaseDsply  Displays the difference between the current
'                                  public code and the pending release code.
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
    
    Dim sComp As String
    
    With p_comp
        sComp = .CompName
        If .CodePnding Is Nothing Then
            Set .CodePnding = New clsCode
            .CodePnding.Source = CommonPending.LastModExpFile(sComp)
        End If
        If .CodeLines Is Nothing Then
            Set .CodeCrrent = New clsCode
            .CodeCrrent.Source = .VBComp.CodeModule
        End If
        PendingVersusServicedCode = .CodePnding.DiffersFrom(.CodeExprtd)
    End With
    
End Function

Public Function PendingVersusServicedCodeBttn(ByVal p_comp As String) As String
    PendingVersusServicedCodeBttn = "Display difference" & vbLf & _
                                    mBasic.Spaced(p_comp) & vbLf & _
                                    "Pending versus current code" & vbLf & _
                                    "(VBComponent.CodeModule)"
End Function

Private Function PendingVersusServicedCodeTitleLeft(ByVal p_comp As String) As String
    PendingVersusServicedCodeTitleLeft = "Pending  " & mBasic.Spaced(p_comp) & "  (Common-Components/PendingReleases folder)"
End Function

Private Function PendingVersusServicedCodeTitleRight(ByVal p_comp As String) As String
    PendingVersusServicedCodeTitleRight = "Current serviced  " & mBasic.Spaced(p_comp) & "  (VBComponent.CodeModule)"
End Function

Public Function PendingVersusServicedExport(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the pending code differs from the serviced Export-File's
' code.
' ----------------------------------------------------------------------------
    
    Dim sComp As String
    
    With p_comp
        sComp = .CompName
        If .CodePnding Is Nothing Then
            Set .CodePnding = New clsCode
            .CodePnding.Source = CommonPending.LastModExpFile(sComp)
        End If
        If .CodeExprtd Is Nothing Then
            Set .CodeExprtd = New clsCode
            .CodeExprtd.Source = .ExpFileFullName
        End If
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
    Dim sComp As String
    
    With p_comp
        sComp = .CompName
        If .CodePnding Is Nothing Then
            Set .CodePnding = New clsCode
            .CodePnding.Source = CommonPending.LastModExpFile(sComp)
        End If
        If .CodeExprtd Is Nothing Then
            Set .CodeExprtd = New clsCode
            .CodeExprtd.Source = .ExpFileFullName
        End If
        .CodePnding.DsplyDiffs d_left_this_file_name:="PendingRelease" _
                             , d_left_this_file_title:=mDiff.PendingVersusServicedExportTitleLeft(p_comp) _
                             , d_right_versus_code:=.CodeExprtd _
                             , d_right_versus_file_name:="LastModExport" _
                             , d_right_versus_file_title:=mDiff.PendingVersusServicedExportTitleRight(p_comp)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function PendingVersusServicedExportTitleLeft(ByVal p_comp As String) As String
    PendingVersusServicedExportTitleLeft = "Pending release  " & mBasic.Spaced(p_comp) & "  (Common-Components/PendingReleases folder)"
End Function

Private Function PendingVersusServicedExportTitleRight(ByVal p_comp As String) As String
    PendingVersusServicedExportTitleRight = "Current serviced  " & mBasic.Spaced(p_comp) & "  (Export-File)"
End Function

Public Function PublicVersusPendingRelease(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the public code differs from the pending code.
' ----------------------------------------------------------------------------
    
    Dim sComp As String
    
    With p_comp
        sComp = .CompName
        If .CodePublic Is Nothing Then
            Set .CodePublic = New clsCode
            .CodePublic.Source = CommonPublic.LastModExpFile(sComp)
        End If
        If .CodePnding Is Nothing Then
            Set .CodePnding = New clsCode
            .CodePnding.Source = CommonPending.LastModExpFile(sComp)
        End If
        PublicVersusPendingRelease = .CodePublic.DiffersFrom(.CodePnding)
    End With
    
End Function

Public Function PublicVersusPendingReleaseBttn(ByVal p_comp As String) As String
    PublicVersusPendingReleaseBttn = "Display difference" & vbLf & _
                                     mBasic.Spaced(p_comp) & vbLf & _
                                     "Public versus Pending"
End Function

Public Sub PublicVersusPendingReleaseDsply(ByVal p_comp As clsComp)
    Const PROC = "PublicVersusPendingReleaseDsply"
    
    On Error GoTo eh
    Dim sComp As String
    
    With p_comp
        sComp = .CompName
        .CodePublic.Source = CommonPublic.LastModExpFile(sComp)
        .CodePnding.Source = CommonPending.LastModExpFile(sComp)
        .CodePublic.DsplyDiffs d_left_this_file_name:="PublicLastModExpFile" _
                             , d_left_this_file_title:=mDiff.PublicVersusPendingReleaseTitleLeft(sComp) _
                             , d_right_versus_code:=.CodePnding _
                             , d_right_versus_file_name:="CodePending" _
                             , d_right_versus_file_title:=mDiff.PublicVersusPendingReleaseTitleRight(sComp)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function PublicVersusPendingReleaseTitleLeft(ByVal p_comp As String) As String
    PublicVersusPendingReleaseTitleLeft = "Public  " & mBasic.Spaced(p_comp) & "  (Common-Components folder)"
End Function

Private Function PublicVersusPendingReleaseTitleRight(ByVal p_comp As String) As String
    PublicVersusPendingReleaseTitleRight = "Pending release " & mBasic.Spaced(p_comp) & " (Common-Components\PendingReleases folder)"
End Function

Public Function ServicedExportVersusPublic(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the serviced compoent's (p_comp) Export-File differs from
' the public one in the Common-Components folder.
' ----------------------------------------------------------------------------
    
    Dim sComp   As String
    Dim arrDiff As Variant
    Dim v       As Variant
    
    With p_comp
        sComp = .CompName
        If .CodePublic Is Nothing Then
            Set .CodePublic = New clsCode
            .CodePublic.Source = CommonPublic.LastModExpFile(sComp)
        End If
        If .CodeExprtd.DiffersFrom(.CodePublic, , arrDiff) Then
            ServicedExportVersusPublic = True
            With Services
                For Each v In arrDiff
                    .Log(sComp) = "Code change: " & v
                Next v
            End With
        End If
    End With
    
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
    Dim sComp As String
    
    With p_comp
        sComp = .CompName
        .CodeExprtd.DsplyDiffs d_left_this_file_name:="CommonComponentServiced" _
                             , d_left_this_file_title:=mDiff.ServicedExportVersusPublicTitleLeft(sComp) _
                             , d_right_versus_code:=.CodePublic _
                             , d_right_versus_file_name:="CommonComponentPublic" _
                             , d_right_versus_file_title:=mDiff.ServicedExportVersusPublicTitleRight(sComp)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ServicedExportVersusPublicTitleLeft(ByVal p_comp As String) As String
    ServicedExportVersusPublicTitleLeft = "Serviced  " & mBasic.Spaced(p_comp) & "  (Export-File)"
End Function

Public Function ServicedExportVersusServicedCode(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the serviced compoent's (p_comp) Export-File differs from
' the public one in the Common-Components folder.
' ----------------------------------------------------------------------------
    
    Dim sComp   As String
    Dim arrDiff As Variant
    Dim v       As Variant
    
    With p_comp
        sComp = .CompName
        If .CodeCrrent.DiffersFrom(.CodeExprtd, , arrDiff) Then
            ServicedExportVersusServicedCode = True
            If .IsCommComp Then
                With Services
                    .Log(sComp) = "Common Component code change: " & v
                    For Each v In arrDiff
                        .Log(sComp) = v
                    Next v
                End With
            End If
        End If
    End With
    
End Function

Private Function ServicedExportVersusPublicTitleRight(ByVal p_comp As String) As String
    ServicedExportVersusPublicTitleRight = "Public " & p_comp & " (Common-Components folder)"
End Function

Public Function PublicVersusServicedExport(ByVal p_comp As clsComp) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the pending code differs from the serviced Export-File's
' code.
' ----------------------------------------------------------------------------
    
    Dim sComp As String
    
    With p_comp
        sComp = .CompName
        If .CodePublic Is Nothing Then
            Set .CodePublic = New clsCode
            .CodePublic.Source = CommonPublic.LastModExpFile(sComp)
        End If
        If .CodeExprtd Is Nothing Then
            Set .CodeExprtd = New clsCode
            .CodeExprtd.Source = .ExpFileFullName
        End If
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
    Dim sComp As String
    
    With p_comp
        sComp = .CompName
        If .CodePublic Is Nothing Then
            Set .CodePublic = New clsCode
            .CodePublic.Source = CommonPublic.LastModExpFile(sComp)
        End If
        If .CodeExprtd Is Nothing Then
            Set .CodeExprtd = New clsCode
            .CodeExprtd.Source = .ExpFileFullName
        End If
        .CodePublic.DsplyDiffs d_left_this_file_name:="LastModExpFile" _
                             , d_left_this_file_title:=mDiff.PublicVersusServicedExportTitleLeft(p_comp) _
                             , d_right_versus_code:=.CodeExprtd _
                             , d_right_versus_file_name:="CurrentExportedModifications" _
                             , d_right_versus_file_title:=mDiff.PublicVersusServicedExportTitleRight(p_comp)
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function PublicVersusServicedExportTitleLeft(ByVal p_comp As String) As String
    PublicVersusServicedExportTitleLeft = "Public  " & mBasic.Spaced(p_comp) & "  (Common-Components folder)"
End Function

Private Function PublicVersusServicedExportTitleRight(ByVal p_comp As String) As String
    PublicVersusServicedExportTitleRight = "Current serviced  " & mBasic.Spaced(p_comp) & "  (Export-File)"
End Function



