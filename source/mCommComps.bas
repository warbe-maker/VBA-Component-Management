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

Public Property Get BttnInconsistencyExport() As String
    BttnInconsistencyExport = "Export" & vbLf & _
                              "(the hosted version is the one up-to-date)"
End Property

Public Property Get BttnInconsistencySkip() As String
    BttnInconsistencySkip = "Skip" & vbLf & "for further investigation"
End Property

Public Property Get BttnInconsistencyUpdate() As String
    BttnInconsistencyUpdate = "Update (re-import)" & vbLf & _
                              "(the 'Common-Components Folder' version is up-to-date)"

End Property

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

Public Function ExistsAsGlobalCommonComponentExportFile(ByVal x_vbc As VBComponent) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the VBComponent's (x_vbc) Export-File exists in the
' global Common-Components-Folder.
' ----------------------------------------------------------------------------
    Const PROC  As String = "ExistsAsGlobalCommonComponentExportFile"
    
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    Dim Comp    As New clsComp
    
    mBasic.BoP ErrSrc(PROC)
    With Comp
        Set .Wrkbk = Services.Serviced
        .CompName = x_vbc.Name
        sFile = wsConfig.FolderCommonComponentsPath & "\" & x_vbc.Name & .ExpFileExt
    End With
    ExistsAsGlobalCommonComponentExportFile = fso.FileExists(sFile)
    Set Comp = Nothing
    
xt: mBasic.EoP ErrSrc(PROC)
    
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

Public Function InconsitencyWarning(ByVal i_file_full_name, _
                                    ByVal i_file_full_name_saved, _
                                    ByVal i_message) As Variant
' ----------------------------------------------------------------------------
' Displays an information about a modification of a Used Common Component.
' The display offers the option to display the code difference. The function
' returns the finally pressed button.
' ----------------------------------------------------------------------------
    Const PROC = "InconsitencyWarning"
    
    On Error GoTo eh
    Dim fso                 As New FileSystemObject
    Dim Msg                 As mMsg.TypeMsg
    Dim cllBttns            As Collection
    Dim BttnDsply           As String
    Dim BttnSkip            As String
    Dim BttnExport          As String
    Dim BttnUpdate          As String
    
    BttnDsply = "Display difference" & "(of Export-Files)"
    BttnSkip = BttnInconsistencySkip
    BttnExport = BttnInconsistencyExport
    BttnUpdate = BttnInconsistencyUpdate
    
    Set cllBttns = mMsg.Buttons(BttnDsply, vbLf, BttnExport, BttnUpdate, vbLf, BttnSkip)
    With Msg.Section(1)
        With .Label
            .Text = "Attention!"
            .FontColor = rgbRed
        End With
        With .Text
            .Text = i_message
        End With
    End With
    With Msg.Section(2)
        .Label.Text = Replace(Replace(BttnExport, vbLf, " "), "  ", " ")
        .Label.FontColor = rgbBlue
        .Text.Text = "According to the displayed differencies is the hosted version the one up-to-date. " & vbLf & _
                     "It will be exported and copied to the 'Common-Components Folder' and the 'Revision Number' will be increased."
    End With
    With Msg.Section(3)
        .Label.Text = Replace(Replace(BttnUpdate, vbLf, " "), "  ", " ")
        .Label.FontColor = rgbBlue
        .Text.Text = "The hosted version will be updated with the 'Common-Components Folder' version (by re-import) and " & _
                     "the 'Revision-Number' of the hosted version will be set identical with the 'Common-Components Folder' version"
    End With
    With Msg.Section(4)
        .Label.Text = Replace(Replace(BttnSkip, vbLf, " "), "  ", " ")
        .Label.FontColor = rgbBlue
        .Text.Text = "Will be clarified later!& " & vbLf & _
                     "Note: Each Workbook Safe will redisplay this message until either Export or Update is performed."
    End With
    With Msg.Section(5)
        .Label.Text = "Background:"
        .Label.FontColor = rgbBlue
        .Text.Text = "When 'hosted Common Component' is modified within its hosting Workbook and exported, the " & _
                     "'Export File' is copied to the 'Common-Components Folder' and the 'Revision Number' is increased " & _
                     "and set equal in the hosting Workbook's and the Common-Component Folder's ""CommComps.dat"" file. " & _
                     "When a 'used Common Component is modified within the VB-Project just using (not hosting!) it, " & _
                     "the 'Revision Number' only of this 'used Common Component' is increased. When by accident both " & _
                     "had been modified the 'Revision Numbers' may be equal but the Export Files will differ."
    End With
        
    Do
        Select Case mMsg.Dsply(dsply_title:="Inconsistency warning for " & fso.GetBaseName(i_file_full_name) & "!" _
                             , dsply_msg:=Msg _
                             , dsply_buttons:=cllBttns _
                              )
            Case BttnDsply
                Services.ExpFilesDiffDisplay e_file_left_full_name:=i_file_full_name _
                                           , e_file_left_title:="Raw Common Component's Export File: (" & i_file_full_name & ")" _
                                           , e_file_right_full_name:=i_file_full_name_saved _
                                           , e_file_right_title:="Saved Raw's Export File (" & i_file_full_name_saved & ")"
            Case BttnSkip:      InconsitencyWarning = BttnSkip:     Exit Do
            Case BttnExport:    InconsitencyWarning = BttnExport:   Exit Do
            Case BttnUpdate:    InconsitencyWarning = BttnUpdate:   Exit Do
        End Select
    Loop

xt: Set fso = Nothing
    Exit Function

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
    Else
        Services.DsplyProgress "used Common Components updated"
        Services.LogEntrySummary Application.StatusBar
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
    Qoutdated.First Comp ' Get the next outdated component from the queue
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
                .FontColor = rgbGreen
            End With
            .Text.Text = "The component, known as a potential 'Used Common Component' will be de-registered " & _
                         "and ignored in the future! I.e. the ""used"" status will be changed into ""private"" status. " & _
                         "Re-instantiating as a ""used"" Common Component will requires the following steps:" & vbLf & _
                         "1. Remove it" & vbLf & _
                         "2. Save the Workbook" & vbLf & _
                         "3. Re-Import it from the ""Common-Components"" folder."
        End With
        If Comp.KindOfComp = enCommCompUsed Then
            If Comp.RevisionNumber > Comp.Raw.RevisionNumber Then
                With .Section(4)
                    .Label.Text = "Attention!!!!"
                    .Text.Text = "The ""Revision Number"" of this ""Used Common Component"" is greater than the ""Revision Number"" " & _
                                 "of the ""Common Component"" in the ""Common Component Folder""." & vbLf & _
                                 "The likely reason: The ""Used Common Component"" has been modified with the Workbook using it. " & _
                                 "When updated, this modification will get lost. Checking the changes and making them in the ""raw/hosted"" " & _
                                 "version is a choice. Another would be to " & Replace(BttnSkipForNow, vbLf, " ") & "." & vbLf & _
                                 "Another, less likely reason may be that the ""Common Components Folder"", only possible when it is " & _
                                 "updated on different computers!"
                End With
            ElseIf Comp.RevisionNumber = Comp.Raw.RevisionNumber Then
                With .Section(4)
                    .Label.Text = "Attention!!!!"
                    .Text.Text = "Though the ""Revision Number"" of this ""Used Common Component"" is  e q u a l  to the ""Revision Number"" " & _
                                 "of the ""Common Component"" in the ""Common Component Folder"" the code differs!" & vbLf & _
                                 "The likely reason: The ""Used Common Component"" and its ""raw"" had botth been inconsistenly! modified. " & _
                                 "When updated, the modification in the ""Used Common Component"" will get lost. Checking the difference should " & _
                                 "indicate whether or not an update is preferred. Skipping the update for now postpones the decision."
                End With
            End If
        End If
    End With
    
    mMsg.BttnAppRun AppRunArgs, BttnUpdate _
                                , ThisWorkbook _
                                , "mCommComps.OutdatedUpdateChoiceUpdate" _
                                , Comp.CompName
    mMsg.BttnAppRun AppRunArgs, BttnDsplyDiffs _
                                , ThisWorkbook _
                                , "mCommComps.OutdatedUpdateChoiceDsplyDiffs"
    mMsg.BttnAppRun AppRunArgs, BttnSkipForNow _
                                , ThisWorkbook _
                                , "mCommComps.OutdatedUpdateChoiceSkipForNow" _
                                , Comp.CompName
    mMsg.BttnAppRun AppRunArgs, BttnSkipForever _
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
        Services.ExpFilesDiffDisplay .ExpFileFullName, .Raw.SavedExpFileFullName, "Currently used (" & .ExpFileFullName & ")", "Up-to-date (" & .Raw.SavedExpFileFullName & ")"
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
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    mCompManDat.RegistrationState(u_comp_name) = enRegStatePrivate
    Qoutdated.DeQueue
    Set wbk = Services.Serviced
    With New clsComp
        Set .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
    End With
    
    With Services
        .NoOfItemsIgnored = .NoOfItemsIgnored + 1
        .LogEntry "Outdated used Commpon Component: Update skipped forever!"
    End With
    
xt: Services.MessageUnload UpdateDialogTitle
    mCommComps.OutdatedUpdate
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdateChoiceSkipForNow(ByVal u_comp_name As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoiceSkipForNow"
    
    On Error GoTo eh
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.DeQueue
    Set wbk = Services.Serviced
    With New clsComp
        Set .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
    End With
    
    With Services
        .NoOfItemsIgnored = .NoOfItemsIgnored + 1
        .LogEntry "Outdated used Commpon Component: Update skipped for now!"
    End With
    
xt: Services.MessageUnload UpdateDialogTitle
    mCommComps.OutdatedUpdate
    mBasic.EoP ErrSrc(PROC)
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
    Dim sFile   As String
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.First Comp
    Set wbk = Services.Serviced
    Set Comp = New clsComp
    With Comp
        Set .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
        v = .KindOfComp
        sFile = .Raw.SavedExpFileFullName
        mUpdate.ByReImport b_wbk_target:=wbk _
                         , b_vbc_name:=u_comp_name _
                         , b_exp_file:=sFile
        With Services
            .NoOfItemsServiced = .NoOfItemsServiced + 1
            .NoOfItemsServicedNames = u_comp_name
            .DsplyProgress "used Common Components updated"
            .LogEntry "Outdated used Common Component updated by re-import of " & sFile
        End With
    End With
    Qoutdated.DeQueue
    Set Comp = Nothing
    
xt: Services.MessageUnload UpdateDialogTitle
    mCommComps.OutdatedUpdate
    mBasic.EoP ErrSrc(PROC)
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
    Dim vbc         As VBComponent
    Dim fso         As New FileSystemObject
    Dim Comp        As clsComp
    Dim wbk         As Workbook
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim sName       As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    Set Qoutdated = New clsQ
    Set dct = Comps.All ' all = all relevant for the current service
    
    For Each v In dct
        Set Comp = dct(v)
        With Comp
            Services.ServicedItem = .VBComp
            If .Outdated Then
                Qoutdated.EnQueue Comp
                sName = .CompName
                With Services
                    .NoOfItemsServicedNames = sName
                    .NoOfItemsOutdated = Qoutdated.Size
                    .DsplyProgress "collected Common Components used outdated"
                End With
            Else
                '~~ When not outdated due to a code difference the revision numbers ought to be equal
                If .RevisionNumber <> .Raw.RevisionNumber Then .RevisionNumber = .Raw.RevisionNumber
                With Services
                    .NoOfItemsIgnored = .NoOfItemsIgnored + 1
                    .LogEntry "Used Common Component is up-to-date"
                End With
            End If ' .Outdated
        End With
        Set Comp = Nothing
        Services.DsplyProgress "collected Common Components used outdated"
    Next v
    Services.DsplyProgress "collected Common Components used outdated"
    
xt: If wsService.CommonComponentsUsed = 0 Then wsService.CommonComponentsUsed = Services.NoOfItemsCommonUsed
    If wsService.CommonComponentsOutdated = 0 Then wsService.CommonComponentsOutdated = Qoutdated.Size
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
    Const PROC  As String = "Register"
    
    Dim fso     As New FileSystemObject
    Dim wbk     As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    RawHostWbFullName(r_comp_name) = wbk.FullName
    RawHostWbName(r_comp_name) = wbk.Name
    RawHostWbBaseName(r_comp_name) = fso.GetBaseName(wbk.FullName)
    RawExpFileFullName(r_comp_name) = r_exp_file
    Set fso = Nothing

xt: mBasic.EoP ErrSrc(PROC)
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
                                        ByVal stgf_file As File, _
                                        ByVal stgf_file_full_name As String)
' ------------------------------------------------------------------------------
' Save a copy of the hosted raw`s (stgf_comp_name) export file to the Common
' Components folder which serves as source for the update of Common Components
' used in other VB-Projects.
' ------------------------------------------------------------------------------
    Const PROC  As String = "SaveToCommonComponentsFolder"
    
    Dim frxFile As File
    Dim fso     As New FileSystemObject
    
    mBasic.BoP ErrSrc(PROC)
    mCommComps.SavedExpFile(stgf_comp_name) = stgf_file
    '~~ When the Export file has a .frm extension the .frx file needs to be copied too
    If fso.GetExtensionName(stgf_file_full_name) = "frm" Then
        Set frxFile = fso.GetFile(Replace(stgf_file_full_name, "frm", "frx"))
        mCommComps.SavedExpFile(stgf_comp_name) = frxFile
    End If

    mCommComps.RawExpFileFullName(stgf_comp_name) = stgf_file_full_name
    mCommComps.RawHostWbBaseName(stgf_comp_name) = fso.GetBaseName(Services.Serviced.FullName)
    mCommComps.RawHostWbFullName(stgf_comp_name) = Services.Serviced.FullName
    mCommComps.RawHostWbName(stgf_comp_name) = Services.Serviced.Name
    mCommComps.RevisionNumber(stgf_comp_name) = mCompManDat.RawRevisionNumber(stgf_comp_name)
    
xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)

End Sub

