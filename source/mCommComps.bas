Attribute VB_Name = "mCommComps"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mCommComps: Management of Common Components
'                             a) in the dedicated Common Components folder
'                             b) in the CommComps.dat
' Public services:
' -
' ---------------------------------------------------------------------------
Public Qoutdated                                As clsQ

Private Const VNAME_RAW_HOST_NAME               As String = "RawHostName"
Private Const VNAME_RAW_HOST_BASE_NAME          As String = "RawHostBaseName"
Private Const VNAME_RAW_HOST_FULL_NAME          As String = "RawHostFullName"
Private Const VNAME_RAW_SAVED_REVISION_NUMBER   As String = "RawRevisionNumber"
Private Const VNAME_RAW_EXP_FILE_FULL_NAME      As String = "RawExpFileFullName"

Private BttnUpdate                              As String
Private BttnDsplyDiffs                          As String
Private BttnSkipForNow                          As String
Private BttnSkipForever                         As String
Private UpdateDialogTitle                       As String
Private UpdateDialogTop                         As Long
Private UpdateDialogLeft                        As Long

Public Property Get CommCompsDatFileFullName() As String
    CommCompsDatFileFullName = wsConfig.FolderCommonComponentsPath & "\CommComps.dat"
End Property

Public Property Get RawExpFileFullName(Optional ByVal comp_name As String) As String
    RawExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME)
    If RawExpFileFullName = vbNullString Then
    End If
End Property

Public Property Let RawExpFileFullName(Optional ByVal comp_name As String, _
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

Public Property Get RevisionNumber(Optional ByVal comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    RevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_SAVED_REVISION_NUMBER)
End Property

Public Property Let RevisionNumber(Optional ByVal comp_name As String, _
                                            ByVal comp_rev_no As String)
' ------------------------------------------------------------------------------
' Returns a revision number in the form yy-mm-dd.00. Plus one when an existing
' revision number is provided (comp_rev-no) or the current date with .01.
' ------------------------------------------------------------------------------
    Const PROC = "RevisionNumber Let"
    
    On Error GoTo eh
    
    If comp_rev_no = vbNullString Then
        Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_SAVED_REVISION_NUMBER) = mCompManDat.RevisionNumberInitial
    Else
        Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_SAVED_REVISION_NUMBER) = comp_rev_no
    End If
    
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

Private Property Let SavedExpFile(Optional ByVal comp_name, _
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
        Case "used":    CommCompRegStateEnum = mComp.enRegStateUsed
        Case "private": CommCompRegStateEnum = mComp.enRegStatePrivate
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

Public Function Hosted(ByVal h_hosted As String) As Dictionary
' ----------------------------------------------------------------------------
' Returns a string of component names (h_hosted) as Dictionary.
' ----------------------------------------------------------------------------
    Dim v       As Variant
    Dim sComp   As String
    Dim dct     As New Dictionary
    
    For Each v In Split(h_hosted, ",")
        sComp = Trim$(v)
        If Not dct.Exists(sComp) Then
            dct.Add sComp, sComp
        End If
    Next v
    Set Hosted = dct
    Set dct = Nothing
    
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
Private Function MaxRawLenght() As Long
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

Public Sub OutdatedUpdate()
' ------------------------------------------------------------------------------
' Collects all used outdated Common Components when called for the first time
' and displays the first one queued in Qoutdated. The service is re-called until
' the queue is empty. The display of the update choices is a mode-less dialog
' which calls sub-services in accordance with the button pressed.
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdate"
    
    On Error GoTo eh
    If Qoutdated Is Nothing Then OutdatedUpdateCollect
    If Not Qoutdated.IsEmpty Then
        OutdatedUpdateChoice
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoice()
' ------------------------------------------------------------------------------
' Displays the first outdated Common Component in the queue Qoutdated in a mode-
' less dialog for one of the options: "update", "display diffs", "skip for now",
' or "skip forever".
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoice"
    
    On Error GoTo eh
    Dim AppRunArgs  As Dictionary
    Dim cllButtons  As Collection
    Dim Comp        As clsComp
    Dim fUpdate     As fMsg
    Dim Msg         As mMsg.TypeMsg

    Set AppRunArgs = New Dictionary
    Qoutdated.First Comp
    BttnUpdate = "Update" & vbLf & vbLf & Comp.CompName
    BttnDsplyDiffs = "Display changes"
    BttnSkipForNow = "Skip this update" & vbLf & "for now"
    BttnSkipForever = "Skip this update" & vbLf & "f o r e v e r" & vbLf & "(I am aware of the consequence)"
    Set cllButtons = mMsg.Buttons(BttnUpdate, vbLf, BttnDsplyDiffs, vbLf, BttnSkipForNow, BttnSkipForever)
    UpdateDialogTitle = "Used Common Component  " & mBasic.Spaced(Comp.CodeName) & "  has changed"
    Set fUpdate = mMsg.MsgInstance(UpdateDialogTitle)
    
    With Msg
        With .Section(1)
            .Text.Text = "Updating the used Common Component  " & mBasic.Spaced(Comp.CompName) & "  is due!"
        End With
        With .Section(2)
            With .Label
                .Text = Replace(BttnSkipForNow, vbLf, " ")
                .FontBold = True
            End With
            .Text.Text = "Update will be proposed with next open"
        End With
        With .Section(3)
            With .Label
                .Text = Replace(BttnSkipForever, vbLf, " ")
                .FontBold = True
                .FontColor = rgbRed
            End With
            .Text.Text = "Although the component's name is identical with a known Common Component " & _
                         "it will be de-registered as a ""used"" one and registered ""private"" instead. " & _
                         "Re-instating the component as a ""Used Common Component"" requires the following steps:" & vbLf & _
                         "1. Rename it" & vbLf & _
                         "2. Save the Workbook" & vbLf & _
                         "3. Import the Common Component from the ""Common-Components"" folder"
        End With
    End With
    
    mMsg.ButtonAppRun AppRunArgs, BttnUpdate _
                                , ThisWorkbook _
                                , "mCommComps.OutdatedUpdateChoiceUpdate" _
                                , Comp.CompName
    mMsg.ButtonAppRun AppRunArgs, BttnDsplyDiffs _
                                , ThisWorkbook _
                                , "mCommComps.OutdatedUpdateChoiceDsplyDiffs"
    mMsg.ButtonAppRun AppRunArgs, BttnSkipForNow _
                                , ThisWorkbook _
                                , "mCommComps.OutdatedUpdateChoiceSkipForNow"
    mMsg.ButtonAppRun AppRunArgs, BttnSkipForever _
                                , ThisWorkbook _
                                , "mCommComps.OutdatedUpdateChoiceSkipForever" _
                                , Comp.CompName
    
    '~~ Display the mode-less dialog for the Names synchronization to run
    mMsg.Dsply dsply_title:=UpdateDialogTitle _
                 , dsply_msg:=Msg _
                 , dsply_buttons:=cllButtons _
                 , dsply_modeless:=True _
                 , dsply_buttons_app_run:=AppRunArgs _
                 , dsply_width_min:=45 _
                 , dsply_pos:=UpdateDialogTop & ";" & UpdateDialogLeft
    DoEvents
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoiceDsplyDiffs()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoiceDsplyDiffs"
    
    On Error GoTo eh
    Dim Comp    As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.First Comp
    With Comp
        mService.ExpFilesDiffDisplay .ExpFileFullName, .Raw.SavedExpFileFullName, "Currently used (" & .ExpFileFullName & ")", "Up-to-date (" & .Raw.SavedExpFileFullName & ")"
    End With
    Set Comp = Nothing
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoiceSkipForever(ByVal u_comp_name)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoiceSkipForever"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mCompManDat.RegistrationState(u_comp_name) = enRegStatePrivate
    Qoutdated.DeQueue
    
xt: mBasic.EoP ErrSrc(PROC)
    mService.MessageUnload UpdateDialogTitle
    mCommComps.OutdatedUpdate
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoiceSkipForNow()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoiceSkipForNow"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.DeQueue
    
xt: mBasic.EoP ErrSrc(PROC)
    mService.MessageUnload UpdateDialogTitle
    mCommComps.OutdatedUpdate
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoiceUpdate(ByVal u_comp_name As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoiceUpdate"
    
    On Error GoTo eh
    Dim wbk     As Workbook
    Dim Comp    As clsComp
    Dim v       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.First Comp
    Set wbk = mService.WbkServiced
    Set Comp = New clsComp
    With Comp
        Set .Wrkbk = wbk
        .CompName = u_comp_name
        v = .KindOfComp
        mUpdate.ByReImport b_wbk_target:=wbk _
                         , b_vbc_name:=u_comp_name _
                         , b_exp_file:=.Raw.SavedExpFileFullName
    End With
    Qoutdated.DeQueue
    Set Comp = Nothing
    
xt: mBasic.EoP ErrSrc(PROC)
    mService.MessageUnload UpdateDialogTitle
    mCommComps.OutdatedUpdate
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateCollect()
' ------------------------------------------------------------------------------
' Collects all outdated Used Common Components and enqueues them in Qoutdated.
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateCollect"
    
    On Error GoTo eh
    Dim dctAll      As Dictionary
    Dim dctOutdated As New Dictionary
    Dim vbc         As VBComponent
    Dim fso         As New FileSystemObject
    Dim sOutdated   As String
    Dim Comp        As clsComp
    Dim lAll        As Long
    Dim lRemaining  As Long
    Dim wbk         As Workbook
    Dim lUsed       As Long
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = mService.WbkServiced
    Set dctAll = mService.AllComps(wbk)
    Set Qoutdated = New clsQ
    Application.StatusBar = vbNullString
    
    With wbk.VBProject
        lAll = .VBComponents.Count
        lRemaining = lAll
        For Each vbc In .VBComponents
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = wbk
                .CompName = vbc.Name
                Set .VBComp = vbc
                If .KindOfComp = mCompMan.enCommCompUsed Then
                    lUsed = lUsed + 1
                    If .Outdated Then
                        Qoutdated.EnQueue Comp
                        sOutdated = .CompName
                    Else
                        If .RevisionNumber <> .Raw.RevisionNumber Then
                            '~~ When not outdated due ti a code difference the revision numbers ought to be equal
                            Debug.Print "Revision-Number used: = " & .RevisionNumber
                            Debug.Print "Revision-Number raw:  = " & .Raw.RevisionNumber
                            .RevisionNumber = .Raw.RevisionNumber
                        End If
                    End If ' .Outdated
                End If ' Used Common Component
            End With
            Set Comp = Nothing
            lRemaining = lRemaining - 1
            mService.DsplyStatus _
            "Collect " & mService.Progress(p_result:=Qoutdated.Size _
                                         , p_of:=lAll _
                                         , p_op:="outdated used Common Components " _
                                         , p_comps:=sOutdated _
                                         , p_dots:=lRemaining _
                                          )
        Next vbc
    End With
    mService.DsplyStatus _
    "Collect " & mService.Progress(p_result:=Qoutdated.Size _
                                 , p_of:=lAll _
                                 , p_op:="outdated used Common Components" _
                                 , p_comps:=sOutdated _
                                  )
    
xt: If wsService.CommonComponentsUsed = 0 Then wsService.CommonComponentsUsed = lUsed
    If wsService.CommonComponentsOutdated = 0 Then wsService.CommonComponentsOutdated = dctOutdated.Count
    Set dctOutdated = Nothing
    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Register(ByVal r_comp_name As String, _
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
    mCommComps.RevisionNumber(stgf_comp_name) = mCompManDat.RawRevisionNumber(stgf_comp_name)
    
    Set fso = Nothing
End Sub

