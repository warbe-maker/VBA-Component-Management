Attribute VB_Name = "mHskpng"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mHskpng." & s
End Function

Public Sub CommComps(ByVal h_hosted As String)
' ------------------------------------------------------------------------------
' Removes obsolete sections which are those neither representing an existing
' VBComponent no another valid section's Name.
' ------------------------------------------------------------------------------
    Const PROC = "CommComps"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    CommCompsRemoveObsoleteComponents h_hosted
    CommCompsAddMissingComponents
    CommCompsHosted h_hosted
    CommCompsNotHosted h_hosted
    CommCompsUsed
    ReorgDatFile CommCompsDatFileFullName
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ReorgDatFile(ByVal r_dat_file As String)
    mFso.PPreorg r_dat_file
End Sub

Private Sub CommCompsAddMissingComponents()
' ------------------------------------------------------------------------------
' Adds for each Common Component's Export-File in the Common-Components folder
' a section to the CommComps.dat when missing.
' Note: A missing section indicates a Common Component of wich the Export-File
'       has been copied manually into the Common-Components folder. In CompMan's
'       definition of a Common Component - which is one hosted in a Workbbok
'       where it is developed, maintained, and tested - a manually added one
'       is an orphan until a Workbook claims hosting it.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsAddMissingComponents"
    
    On Error GoTo eh
    Dim fle         As File
    Dim fso         As New FileSystemObject
    Dim dct         As Dictionary
    Dim sCompName   As String
    Dim sExt        As String
    
    Set dct = mCommComps.Components
    With fso
        For Each fle In .GetFolder(wsConfig.FolderCommonComponentsPath).Files
            sExt = .GetExtensionName(fle.Path)
            Select Case sExt
                Case "bas", "frm", "cls"
                    sCompName = .GetBaseName(fle.Path)
                    If Not dct.Exists(sCompName) Then
                        mCommComps.RawExpFileFullName(sCompName) = vbNullString
                        mCommComps.RawHostWbBaseName(sCompName) = vbNullString
                        mCommComps.RawHostWbFullName(sCompName) = vbNullString
                        mCommComps.RawHostWbName(sCompName) = vbNullString
                        mCommComps.RevisionNumber(sCompName) = mCompManDat.RevisionNumberInitial
                    End If
            End Select
        Next fle
    End With
    
    Set fso = Nothing
    Set dct = Nothing

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsRemoveObsoleteComponents(ByVal h_hosted As String)
' ------------------------------------------------------------------------------
' Remove in the PrivateProfile file CommComps.dat:
' - Sections representing VBComponents for which an Export-File does not exist
'   in the Common-Components folder
' - Sections indicating a Common Component of the serviced Workbook but the
'   component not exists.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsRemoveObsoleteComponents"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wbk         As Workbook
    Dim dct         As Dictionary
    Dim fso         As New FileSystemObject
    Dim sBaseName   As String
    Dim dctHosted   As Dictionary
    Dim sExpFile    As String
    
    Set dctHosted = Hosted(h_hosted)
    Set wbk = Services.Serviced
    sBaseName = fso.GetBaseName(wbk.FullName)
    Set dct = mCommComps.Components
    
    '~~ Obsolete because the component is no longer hosted by the indicated Workbook
    '~~ no longer exist in the indicated Workbook
    For Each v In dct
        If mCommComps.RawHostWbBaseName(v) = sBaseName Then
            '~~ The component indicates being one of the serviced Workbook
            If Not mComp.Exists(v, wbk) _
            Or Not dctHosted.Exists(v) Then
                mCompManDat.RemoveComponent v
            End If
        End If
    Next v
    
    '~~ Obsolete because the corresponding Export-File
    '~~ no longer exists in the Common-Components folder
    '~~ De-register global Common Components no longer hosted
    Set dct = mCommComps.Components
    For Each v In dct
        sExpFile = fso.GetFileName(mCommComps.RawExpFileFullName(v))
        If Not fso.FileExists(wsConfig.FolderCommonComponentsPath & "\" & sExpFile) Then
            CommCompsRemoveSection v
        End If
    Next v
    Set dct = mCommComps.Components
    
xt: Set dct = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsHostedClear(ByVal c_comp_name As String)
' ----------------------------------------------------------------------------
' Performed when a Common Component is not/no longer claimed hosted by the
' originally claiming Workbook. The Common Component just becomes an
' orphan but remains a Common Component which is now just one used by the
' Workbook provided it (still) exist in it.
' ----------------------------------------------------------------------------
    Dim wbk As Workbook
    
    Set wbk = Services.Serviced
    
    mCommComps.RawExpFileFullName(c_comp_name) = vbNullString
    mCommComps.RawHostWbBaseName(c_comp_name) = vbNullString
    mCommComps.RawHostWbFullName(c_comp_name) = vbNullString
    mCommComps.RawHostWbName(c_comp_name) = vbNullString

    If mComp.Exists(c_comp_name, wbk) Then
        mCompManDat.RegistrationState(c_comp_name) = enRegStateUsed
    End If

End Sub

Private Sub CommCompsUsed()
' ----------------------------------------------------------------------------
' Manages the registration of used Common Components, done before change
' components are exported and used changed Common Components are updated.
' When not yet registered a confirmation dialog ensures a component not just
' accidentially has the same name. The type of confirmation is registered
' either as "used" or "private" together with the current revision number.
' When none is available the current date is registered on the fly.
' its Revision-Number is
' ----------------------------------------------------------------------------
    Dim vbc             As VBComponent
    Dim wbk             As Workbook
    Dim BttnConfirmed   As String
    Dim BttnPrivate     As String
    Dim Msg             As mMsg.TypeMsg
    
    BttnConfirmed = "Yes!" & vbLf & _
                    "This is a used Common Component" & vbLf & _
                    "identical with the corresponding" & vbLf & _
                    "VBComponent's Export-File in the" & vbLf & _
                    "Common Components folder"
    BttnPrivate = "No!" & vbLf & _
                  "This is a VBComponent which" & vbLf & _
                  "accidentially has the same name."
    
    Set wbk = Services.Serviced
    For Each vbc In wbk.VBProject.VBComponents
        If mCommComps.ExistsAsGlobalCommonComponentExportFile(vbc) Then
            If mCommComps.RevisionNumber(vbc.Name) = vbNullString Then
                mCommComps.RevisionNumber(vbc.Name) = vbNullString
            End If
            If Not mCompManDat.RegistrationState(vbc.Name) = enRegStatePrivate _
            And Not mCompManDat.RegistrationState(vbc.Name) = enRegStateUsed _
            And Not mCompManDat.RegistrationState(vbc.Name) = enRegStateHosted _
            Then
                '~~ Once an equally named VBComponent is registered a private it will no longer be regarded as "used" and updated.
                Msg.Section(1).Text.Text = "The VBComponent named  '" & vbc.Name & "'  is known as a Common Component " & _
                                           "because it exists in the Common Components folder  '" & _
                                           wsConfig.FolderCommonComponentsPath & "'  but is yet not registered either " & _
                                           "as used or private in the serviced Workbook."
                
                Select Case mMsg.Dsply(dsply_title:="Not yet registered Common Component" _
                                     , dsply_msg:=Msg _
                                     , dsply_buttons:=mMsg.Buttons(BttnConfirmed, vbLf, BttnPrivate))
                    Case BttnConfirmed: mCompManDat.RegistrationState(vbc.Name) = enRegStateUsed
                                        mCompManDat.RawRevisionNumber(vbc.Name) = mCommComps.RevisionNumber(vbc.Name)
                    Case BttnPrivate:   mCompManDat.RegistrationState(vbc.Name) = enRegStatePrivate
                End Select
            End If
        Else
            '~~ The Export-File has manually been copied into the Common
            '~~ Components-Folder and thus is yet not registered
            
        End If
    Next vbc

End Sub

Private Sub CommCompsRemoveSection(ByVal s As String)
    mFso.PPremoveSections CommCompsDatFileFullName, s
End Sub

Private Sub CommCompsHosted(ByVal m_hosted As String)
' ----------------------------------------------------------------------------
' - Registers the Workbook as 'Raw-Host' when it hosts at least one Common
'   Component
' - Maintains an up-to-date copy of the Export-File in the Common-Components
'   folder
' - Maintains for each hosted (raw) Common Component the properties:
'   - in the local CommComps.dat:
'     - Component Name
'     - Revision Number
'   - in the ComComps-RawsSaved.dat in the Common Components folder:
'     - Component Name
'     - Export File Full Name
'     - Host Base Name
'     - Host Full Name
'     - Host Name
'     - Revision Number
' ----------------------------------------------------------------------------
    Const PROC = "CommCompsHosted"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim fso             As New FileSystemObject
    Dim RawComComp      As clsComp
    Dim sHostBaseName   As String
    Dim dctHosted       As Dictionary
    Dim wbk             As Workbook
    
    Set wbk = Services.Serviced
    sHostBaseName = fso.GetBaseName(wbk.FullName)
    Set dctHosted = mCommComps.Hosted(m_hosted)
                    
    For Each v In dctHosted
        If Not mComp.Exists(v, wbk) Then
            MsgBox "The VBComponent " & v & " is claimed hosted by the serviced Workbook " & Services.Serviced.Name & _
                   " but is not known by the Workbook's VB-Project - and thus will be ignored!" & vbLf & vbLf & _
                   "When the component is no longer hosted or its name has changed the argument needs to be updated accordingly.", _
                   vbOK, "VBComponent " & v & "unkonwn!"
        Else
            Set RawComComp = New clsComp
            With RawComComp
                Set .Wrkbk = Services.Serviced
                .CompName = v
                If mCompManDat.RegistrationState(v) <> enRegStateHosted Then
                    '~~ Yet not registered as "hosted" as the serviced Workbook claims it
                    mCompManDat.RegistrationState(v) = enRegStateHosted
                    mCompManDat.RawExpFileFullName(v) = .ExpFileFullName    ' in any case update the Export File name
                    .RevisionNumberIncrease                 ' this will initially set it
                    mCommComps.SaveToCommonComponentsFolder v, .ExpFile, .ExpFileFullName
                End If
                
                mCommComps.Register v, .ExpFileFullName
                If Not mCommComps.ExistsAsGlobalCommonComponentExportFile(.VBComp) Then
                    mCommComps.SaveToCommonComponentsFolder v, .ExpFile, .ExpFileFullName
                    
                End If
                
                If Services.FilesDiffer(fd_exp_file_1:=.ExpFile _
                                      , fd_exp_file_2:=mCommComps.SavedExpFile(v)) Then
                    '~~ Attention! This is a cruical issue which should never be the case. However, when different
                    '~~ computers/users are involved in the development process ...
                    '~~ Instead of simply updating the saved raw Export File better have carefully checked the case
                    If mCommComps.RevisionNumber(v) = mCompManDat.RawRevisionNumber(v) Then
                        If InconsitencyWarning _
                           (exp_file_full_name:=.ExpFile.Path _
                          , saved_exp_file_full_name:=mCommComps.SavedExpFile(v).Path _
                          , sri_diff_message:="While the Revision Number of the 'Hosted Raw'  " & mBasic.Spaced(v) & "  is identical with the " & _
                                              "'Saved Raw' their Export Files are different. Compared were:" & vbLf & _
                                              "Hosted Raw Export File = " & .ExpFile.Path & vbLf & _
                                              "Saved Raw Export File  = " & mCommComps.SavedExpFile(v).Path & vbLf & _
                                              "whereby any empty code lines and case differences had been ignored. " & _
                                              "The difference thus really matters!" _
                           ) Then
                            mCommComps.SaveToCommonComponentsFolder v, .ExpFile, .ExpFileFullName
                        End If
                    ElseIf mCommComps.RevisionNumber(v) <> mCompManDat.RawRevisionNumber(v) Then
                        If mCommComps.InconsitencyWarning _
                           (exp_file_full_name:=.ExpFile.Path _
                          , saved_exp_file_full_name:=mCommComps.SavedExpFile(v).Path _
                          , sri_diff_message:="The 'Revision Number' of the 'Hosted Raw Common Component's Export File' and the " & _
                                              "the 'Saved Raw's Export File' differ:" & vbLf & _
                                              "Hosted Raw = " & mCompManDat.RawRevisionNumber(v) & vbLf & _
                                              "Saved Raw  = " & mCommComps.RevisionNumber(v) & vbLf & _
                                              "and also the Export Files differ. Compared were:" & vbLf & _
                                              "Hosted Raw = " & .ExpFile.Path & vbLf & _
                                              "Saved Raw  = " & mCommComps.SavedExpFile(v).Path & vbLf & _
                                              "whereby any empty code lines and case differences had been ignored. " & _
                                              "The difference thus really matters! Updating is not at all " & _
                                              "recommendable before the issue had been clarified." _
                           ) Then
                            mCommComps.SaveToCommonComponentsFolder v, .ExpFile, .ExpFileFullName
                        End If
                    End If
                End If
            End With
            Set RawComComp = Nothing
        End If
    Next v

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsNotHosted(ByVal h_hosted As String)
' ----------------------------------------------------------------------------
' Removes the hosting Workbook for any Common Component when it is no longer
' claimed hosted by the registered Workbook.
' Note: The Common Component remains being an orphan until it is again claimed
'       being hosted by a Workbook.
' ----------------------------------------------------------------------------
    Dim dctHosted   As Dictionary
    Dim dctComps    As Dictionary
    Dim v           As Variant
    Dim wbk         As Workbook
    
    Set wbk = Services.Serviced
    Set dctHosted = mCommComps.Hosted(h_hosted)
    Set dctComps = mCommComps.Components
    For Each v In dctComps
        If RawHostWbName(v) = wbk.Name Then
            If Not dctHosted.Exists(v) Then
                CommCompsHostedClear v
            End If
        End If
    Next v

End Sub



