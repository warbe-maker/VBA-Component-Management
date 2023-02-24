Attribute VB_Name = "mCommComps"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mCommComps: Management of Common Components
'                             a) in the dedicated Common Components folder
'                             b) in the CommComps.dat
' Public services:
' -
' ---------------------------------------------------------------------------
Private Const VNAME_RAW_HOST_NAME               As String = "RawHostName"
Private Const VNAME_RAW_HOST_BASE_NAME          As String = "RawHostBaseName"
Private Const VNAME_RAW_HOST_FULL_NAME          As String = "RawHostFullName"
Private Const VNAME_RAW_SAVED_REVISION_NUMBER   As String = "RawRevisionNumber"
Private Const VNAME_RAW_EXP_FILE_FULL_NAME      As String = "RawExpFileFullName"

Private Property Get CommCompsDatFileFullName() As String
    CommCompsDatFileFullName = wsConfig.FolderCommonComponentsPath & "\CommComps.dat"
End Property

Private Property Get RawExpFileFullName(Optional ByVal comp_name As String) As String
    RawExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME)
    If RawExpFileFullName = vbNullString Then
    End If
End Property

Private Property Let RawExpFileFullName(Optional ByVal comp_name As String, _
                                                 ByVal exp_file_full_name As String)
' ---------------------------------------------------------------------------
' Sets the Export-File-Full-Name based on the provided (exp_file_full_name)
' from which only the File Name is used.
' ---------------------------------------------------------------------------
    With New FileSystemObject
        Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME) = exp_file_full_name
    End With
End Property

Public Property Get RawHostWbBaseName(Optional ByVal comp_name As String) As String
    RawHostWbBaseName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_BASE_NAME)
End Property

Public Property Let RawHostWbBaseName(Optional ByVal comp_name As String, _
                                            ByVal host_wbk_base_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_BASE_NAME) = host_wbk_base_name
End Property

Public Property Get RawHostWbFullName(Optional ByVal comp_name As String) As String
    RawHostWbFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_FULL_NAME)
End Property

Public Property Let RawHostWbFullName(Optional ByVal comp_name As String, _
                                               ByVal hst_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_FULL_NAME) = hst_full_name
End Property

Public Property Get RawHostWbName(Optional ByVal comp_name As String) As String
    RawHostWbName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_NAME)
End Property

Public Property Let RawHostWbName(Optional ByVal comp_name As String, _
                                        ByVal host_wbk_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_HOST_NAME) = host_wbk_name
End Property

Public Property Get RawSavedRevisionNumber(Optional ByVal comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    RawSavedRevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_SAVED_REVISION_NUMBER)
End Property

Public Property Let RawSavedRevisionNumber(Optional ByVal comp_name As String, _
                                                    ByVal comp_rev_no As String)
' ------------------------------------------------------------------------------
' Returns a revision number in the form yy-mm-dd.00. Plus one when an existing
' revision number is provided (comp_rev-no) or the current date with .01.
' ------------------------------------------------------------------------------
                              
    Const PROC = "RawSavedRevisionNumber Let"
    On Error GoTo eh
    Dim RevDate As String
    Dim RevNo   As Long
    
    If comp_rev_no = vbNullString Then
        RevDate = Format(Now(), "yy-mm-dd")
        RevNo = 1
    Else
        RevDate = Split(comp_rev_no, ".")(0)
        RevNo = Split(comp_rev_no, ".")(1)
    End If
    
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_SAVED_REVISION_NUMBER) = RevDate & "." & Format(RevNo, "000")

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get SavedExpFile(Optional ByVal comp_name) As File
    Const PROC = "SavedExpFile Get"
    
    On Error GoTo eh
    Dim FileName    As String
    Dim sPath       As String
    
    sPath = wsConfig.FolderCommonComponentsPath & "\"
    With New FileSystemObject
        FileName = .GetFileName(RawExpFileFullName(comp_name))
        If FileName <> vbNullString Then
            If .FileExists(sPath & FileName) Then
                Set SavedExpFile = .GetFile(sPath & FileName)
            Else
                Set SavedExpFile = .CreateTextFile(sPath & FileName)
            End If
        End If
    End With
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let SavedExpFile(Optional ByVal comp_name, _
                                          ByVal comp_exp_file As File)
' ---------------------------------------------------------------------------
' Copies the Raw Export File from its host Workbook location to the Common
' Components Folder from where it is used as the source for the update of
' outdated Used Common Components.
' ---------------------------------------------------------------------------
    comp_name = comp_name ' not used
    comp_exp_file.Copy Destination:=wsConfig.FolderCommonComponentsPath & "\" & comp_exp_file.Name, OverWriteFiles:=True
End Property

Public Property Get SavedExpFileFullName(Optional ByVal comp_name As String) As String
' ---------------------------------------------------------------------------
' Returns the Export File Full Name which refers to the Export File saved in
' the Common Components folder.
' ---------------------------------------------------------------------------
    With New FileSystemObject
        SavedExpFileFullName = wsConfig.FolderCommonComponentsPath & "\" & .GetFileName(RawExpFileFullName(comp_name))
    End With
End Property

Private Property Get Value(Optional ByVal pp_section As String, _
                           Optional ByVal pp_value_name As String) As Variant
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFso.PPvalue(pp_file:=CommCompsDatFileFullName _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let Value(Optional ByVal pp_section As String, _
                           Optional ByVal pp_value_name As String, _
                                    ByVal pp_value As Variant)
' ------------------------------------------------------------------------------
' Write the value (pp_value) named (pp_value_name) into the file
' CommCompsDatFileFullName.
' ------------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFso.PPvalue(pp_file:=CommCompsDatFileFullName _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Function CommCompRegStateEnum(ByVal s As String) As enCommCompRegState
    Select Case s
        Case "hosted":  CommCompRegStateEnum = enRegStateHosted
        Case "used":    CommCompRegStateEnum = enRegStateUsed
        Case "private": CommCompRegStateEnum = enRegStatePrivate
    End Select
End Function

Public Function CommCompRegStateString(ByVal en As enCommCompRegState) As String
    Select Case en
        Case enRegStateHosted:  CommCompRegStateString = "hosted"
        Case enRegStateUsed:    CommCompRegStateString = "used"
        Case enRegStatePrivate: CommCompRegStateString = "private"
    End Select
End Function

Public Function Components() As Dictionary
    Set Components = mFso.PPsectionNames(CommCompsDatFileFullName)
End Function

Public Sub DeRegisterNoLongerExisting(ByVal d_hosted As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim v As Variant
    Dim fso As New FileSystemObject
    Dim vbc As VBComponent
    Dim wbk As Workbook
    
    Set wbk = mService.WbkServiced
    HostedRaws = d_hosted

    '~~ De-register global Common Components no longer hosted
    For Each v In mCommComps.Components
        If mCommComps.RawHostWbBaseName = fso.GetBaseName(mService.WbkServiced.FullName) Then
            If Not HostedRaws.Exists(v) Then
                '~~ This VBComponent is no longer hosted (it may have been renamed)
                mCommComps.Remove v
            End If
            If Not mComp.Exists(v, wbk) Then
                mCommComps.Remove v
            End If
        End If
    Next v

    '~~ De-register components global Common Components no longer hosted
    For Each v In mCompManDat.Components
        If mCompManDat.CompIsRegistered(v, enRegStateHosted) Then
            If Not HostedRaws.Exists(v) Then
                '~~ This VBComponent is no longer hosted (it may have been renamed)
                mCompManDat.RemoveComponent v
            End If
            If Not mComp.Exists(v, wbk) Then
                mCompManDat.RemoveComponent v
            End If
        End If
    Next v
    Set fso = Nothing
    
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCommComps" & "." & sProc
End Function

Public Function ExistsAsGlobalCommonComponentExportFile(ByVal ex_vbc As VBComponent) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the VBComponent's (ex_vbc) Export-File exists in the
' global Common-Components-Folder.
' ----------------------------------------------------------------------------
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    Dim Comp    As New clsComp
    
    With Comp
        Set .Wrkbk = mService.WbkServiced
        .CompName = ex_vbc.Name
        sFile = wsConfig.FolderCommonComponentsPath & "\" & ex_vbc.Name & .ExpFileExt
    End With
    ExistsAsGlobalCommonComponentExportFile = fso.FileExists(sFile)
    Set Comp = Nothing
    
End Function

Public Function ExistsRegistered(ByVal raw_vbc_name As String) As Boolean
    ExistsRegistered = Components.Exists(raw_vbc_name)
End Function

Public Function InconsitencyWarning(ByVal exp_file_full_name, _
                                    ByVal saved_exp_file_full_name, _
                                    ByVal sri_diff_message) As Boolean
' ----------------------------------------------------------------------------
' Displays an information about a modification of a Used Common Component.
' The disaplay offers the option to display the code difference.
' Returns TRUE only when the reply is "go ahead, update anyway"
' ----------------------------------------------------------------------------
    Const PROC = "InconsitencyWarning"
    
    On Error GoTo eh
    Dim Msg         As mMsg.TypeMsg
    Dim cllBttns    As Collection
    Dim BttnDsply   As String
    Dim BttnSkip    As String
    Dim BttnAnyway  As String
    
    BttnDsply = "Display code difference" & vbLf & "between hosted and saved" & vbLf & "Export Files"
    BttnSkip = "Do not update!" & vbLf & "further investigation" & vbLf & "is required"
    BttnAnyway = "I know the reason!" & vbLf & "go ahead updating" & vbLf & "(not recommended!)"
    
    Set cllBttns = mMsg.Buttons(BttnDsply, vbLf, BttnAnyway, vbLf, BttnSkip)
    With Msg.Section(1)
        With .Label
            .Text = "Attention!"
            .FontColor = rgbRed
        End With
        With .Text
            .Text = sri_diff_message
            .FontColor = rgbRed
        End With
    End With
    With Msg.Section(2)
        .Label.Text = "Background:"
        .Text.Text = "When a Raw Common Component is modified within its hosting Workbook it is not only exported. " & _
                     "Its 'Revision Number' is increased and the 'Export File' is copied into the 'Common Components' " & _
                     "folder while the 'Revision Number' is updated in the 'ComComps-RawsSaved.dat' file in the " & _
                     "'Common Components' folder. Thus, the Raw's Export File and the copy of it as the 'Revision Number' " & _
                     "are always identical. In case not, something is seriously corrupted."
    End With
        
    Do
        If Not mMsg.IsValidMsgButtonsArg(cllBttns) Then Stop
        Select Case mMsg.Dsply(dsply_title:="Serious inconsistency warning!" _
                             , dsply_msg:=Msg _
                             , dsply_buttons:=cllBttns _
                              )
            Case BttnDsply
                mService.ExpFilesDiffDisplay fd_exp_file_left_full_name:=exp_file_full_name _
                                               , fd_exp_file_left_title:="Raw Common Component's Export File: (" & exp_file_full_name & ")" _
                                               , fd_exp_file_right_full_name:=saved_exp_file_full_name _
                                               , fd_exp_file_right_title:="Saved Raw's Export File (" & saved_exp_file_full_name & ")"
            Case BttnSkip:      InconsitencyWarning = False:    Exit Do
            Case BttnAnyway:    InconsitencyWarning = True:     Exit Do
        End Select
    Loop

xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub ManageHostedCommonComponents(ByVal mh_hosted As String)
' ----------------------------------------------------------------------------
' - Registers the Workbook as 'Raw-Host' when it hosts at least one Common
'   Component
' - Maintaining an up-to-date copy of the Exportfile in a global Common
'   Components Folder.
' - Maintaining for each hosted 'Raw Common Component' the properties:
'   - in the local ComCompsHosted.dat:
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
    Const PROC = "ManageHostedCommonComponents"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim fso             As New FileSystemObject
    Dim RawComComp      As clsComp
    Dim sHostBaseName   As String
    
    mBasic.BoP ErrSrc(PROC)
    sHostBaseName = fso.GetBaseName(mService.WbkServiced.FullName)
    HostedRaws = mh_hosted
                    
    If HostedRaws.Count <> 0 Then
        For Each v In HostedRaws
            Set RawComComp = New clsComp
            With RawComComp
                Set .Wrkbk = mService.WbkServiced
                .CompName = v
                If Not mCompManDat.CompIsRegistered(v, enRegStateHosted) Then
                    '~~ Yet not registered as "hosted" as the serviced Workbook claims it
                    mCompManDat.Register v, enRegStateHosted
                    mCompManDat.RawExpFileFullName(v) = .ExpFileFullName    ' in any case update the Export File name
                    mCompManDat.RawRevisionNumberIncrease v                 ' this will initially set it
                    mCommComps.SaveToCommonComponentsFolder v, .ExpFile, .ExpFileFullName
                End If
                
                mCommComps.Register v, .ExpFileFullName
                If Not mCommComps.ExistsAsGlobalCommonComponentExportFile(.VBComp) Then
                    mCommComps.SaveToCommonComponentsFolder v, .ExpFile, .ExpFileFullName
                    
                End If
                
                If mService.FilesDiffer(fd_exp_file_1:=.ExpFile _
                                      , fd_exp_file_2:=mCommComps.SavedExpFile(v)) Then
                    '~~ Attention! This is a cruical issue which should never be the case. However, when different
                    '~~ computers/users are involved in the development process ...
                    '~~ Instead of simply updating the saved raw Export File better have carefully checked the case
                    If mCommComps.RawSavedRevisionNumber(v) = mCompManDat.RawRevisionNumber(v) Then
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
                    ElseIf mCommComps.RawSavedRevisionNumber(v) <> mCompManDat.RawRevisionNumber(v) Then
                        If mCommComps.InconsitencyWarning _
                           (exp_file_full_name:=.ExpFile.Path _
                          , saved_exp_file_full_name:=mCommComps.SavedExpFile(v).Path _
                          , sri_diff_message:="The 'Revision Number' of the 'Hosted Raw Common Component's Export File' and the " & _
                                              "the 'Saved Raw's Export File' differ:" & vbLf & _
                                              "Hosted Raw = " & mCompManDat.RawRevisionNumber(v) & vbLf & _
                                              "Saved Raw  = " & mCommComps.RawSavedRevisionNumber(v) & vbLf & _
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
        Next v
    End If

xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Register(ByVal r_comp_name As String, _
                     ByVal r_exp_file As String)
    Dim fso As New FileSystemObject
    Dim wbk As Workbook
    
    Set wbk = mService.WbkServiced
    RawHostWbFullName(r_comp_name) = wbk.FullName
    RawHostWbName(r_comp_name) = wbk.Name
    RawHostWbBaseName(r_comp_name) = fso.GetBaseName(wbk.FullName)
    RawExpFileFullName(r_comp_name) = r_exp_file
    Set fso = Nothing

End Sub
Public Sub ManageUsedCommonComponents()
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
    
    Set wbk = mService.WbkServiced
    For Each vbc In wbk.VBProject.VBComponents
        If mCommComps.ExistsAsGlobalCommonComponentExportFile(vbc) Then
            If mCommComps.RawSavedRevisionNumber(vbc.Name) = vbNullString Then
                mCommComps.RawSavedRevisionNumber(vbc.Name) = vbNullString
            End If
            If Not mCompManDat.CompIsRegistered(vbc.Name, enRegStatePrivate) _
            And Not mCompManDat.CompIsRegistered(vbc.Name, enRegStateUsed) _
            And Not mCompManDat.CompIsRegistered(vbc.Name, enRegStateHosted) _
            Then
                '~~ Once an equally named VBComponent is registered a private it will no longer be regarded as "used" and updated.
                Msg.Section(1).Text.Text = "The VBComponent named  '" & vbc.Name & "'  is known as a Common Component " & _
                                           "because it exists in the Common Components folder  '" & _
                                           wsConfig.FolderCommonComponentsPath & "'  but is yet not registered either " & _
                                           "as used or private in the serviced Workbook."
                
                Select Case mMsg.Dsply(dsply_title:="Not yet registered Common Component" _
                                     , dsply_msg:=Msg _
                                     , dsply_buttons:=mMsg.Buttons(BttnConfirmed, vbLf, BttnPrivate))
                    Case BttnConfirmed: mCompManDat.Register vbc.Name, enRegStateUsed
                                        mCompManDat.RawRevisionNumber(vbc.Name) = mCommComps.RawSavedRevisionNumber(vbc.Name)
                    Case BttnPrivate:   mCompManDat.Register vbc.Name, enRegStatePrivate
                End Select
            End If
        Else
            '~~ The Export-File has manually been copied into the Common
            '~~ Components-Folder and thus is yet not registered
            
        End If
    Next vbc

End Sub

Public Function MaxRawLenght() As Long
' -----------------------------------------------
' Returns the max length of a raw componen's name
' -----------------------------------------------
    Const PROC = "MaxRawLenght"
    
    On Error GoTo eh
    Dim v As Variant
    
    For Each v In Components
        MaxRawLenght = Max(MaxRawLenght, Len(v))
    Next v
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Remove(ByVal comp_name As String)
    mFso.PPremoveSections pp_file:=CommCompsDatFileFullName _
                                  , pp_sections:=comp_name
End Sub

Public Function SavedExpFileExists(ByVal comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when an Export File with the name extracted from the
' RawExpFileFullName exists in the Common Components Folder
' ----------------------------------------------------------------------------
    Dim FileName    As String
    With New FileSystemObject
        FileName = .GetFileName(RawExpFileFullName(comp_name))
        SavedExpFileExists = .FileExists(wsConfig.FolderCommonComponentsPath & "\" & FileName)
    End With
End Function

Public Sub SaveToCommonComponentsFolder(ByVal stgf_comp_name As String, _
                                        ByVal stgf_exp_file As File, _
                                        ByVal stgf_exp_file_full_name As String)
' ------------------------------------------------------------------------------
' Save a copy of the hosted raw`s (stgf_comp_name) export file to the Common
' Components folder which serves as source for the update of Common Components
' used in other VB-Projects.
' ------------------------------------------------------------------------------
    Dim frxFile As File
    Dim fso     As New FileSystemObject
    
    mCommComps.SavedExpFile(stgf_comp_name) = stgf_exp_file
    '~~ When the Export file has a .frm extension the .frx file needs to be copied too
    If fso.GetExtensionName(stgf_exp_file_full_name) = "frm" Then
        Set frxFile = fso.GetFile(Replace(stgf_exp_file_full_name, "frm", "frx"))
        mCommComps.SavedExpFile(stgf_comp_name) = frxFile
    End If

    mCommComps.RawExpFileFullName(stgf_comp_name) = stgf_exp_file_full_name
    mCommComps.RawHostWbBaseName(stgf_comp_name) = fso.GetBaseName(mService.WbkServiced.FullName)
    mCommComps.RawHostWbFullName(stgf_comp_name) = mService.WbkServiced.FullName
    mCommComps.RawHostWbName(stgf_comp_name) = mService.WbkServiced.Name
    mCommComps.RawSavedRevisionNumber(stgf_comp_name) = mCompManDat.RawRevisionNumber(stgf_comp_name)
    
    Set fso = Nothing
End Sub

Private Sub Hskpng(ByVal h_hosted As Dictionary)
' ------------------------------------------------------------------------------
' Remove sections in the CommComps.dat referring to VBComponents no longer
' hosted by the serviced Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim v           As Variant
    Dim wbk         As Workbook
    Dim sBaseName   As String
    
    Set wbk = mService.WbkServiced
    sBaseName = fso.GetBaseName(wbk.FullName)
    
    For Each v In mCommComps.Components
        If mCommComps.RawHostWbBaseName(v) = sBaseName Then
            If Not mComp.Exists(v, wbk) _
            Or Not h_hosted.Exists(v) Then
                mCompManDat.RemoveComponent v
            End If
        End If
    Next v
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

