Attribute VB_Name = "mCommComps"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mCommComps: Management of Common Components
'                             a) in the dedicated Common-Components folder
'                             b) in the CommComps.dat
' Public services:
' ----------------
' BttnInconsistencyExport           .
' BttnInconsistencySkip             .
' BttnInconsistencyUpdate           .
' CommCompsDatFileFullName          .
' LastModExpFile               r/w  .
' LastModExpFileFullName       r    .
' LastModExpFileFullNameOrigin r/w  .
' LastModWbk                     w  .
' LastModWbkBaseName           r    .
' LastModWbkFullName           r/w  .
' LastModWbkName               r    .
' RevisionNumber               r/w  .
' ---------------------------------------------------------------------------
Public Qoutdated                                As clsQ

' Housekeeping syntax: (allows the on-the-fly-change of value names as well as their removal)
' "<current-name>:<old-name>" = rename in all sections
' "<current>"                 = no action
' ":<remove-name>"            = remove in all sections
Private Const LAST_MOD_EXP_FILE_FULL_NAME           As String = "LastModExpFileFullName"
Private Const LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN    As String = "LastModExpFileFullNameOrigin"
Private Const LAST_MOD_WBK_FULL_NAME                As String = "LastModWbkFullName"
Private Const REVISION_NUMBER                       As String = "RevisionNumber"

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
                              "(the ""Common-Components Folder"" version is up-to-date)"

End Property

Public Property Get CommCompsDatFileFullName() As String
    CommCompsDatFileFullName = wsConfig.FolderCommonComponentsPath & "\CommComps.dat"
End Property

Public Property Get LastModExpFileFullNameOrigin(Optional ByVal comp_name As String) As String
    LastModExpFileFullNameOrigin = Value(pp_section:=comp_name, pp_value_name:=CurrentName("LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN"))
End Property

Public Property Let LastModExpFileFullNameOrigin(Optional ByVal comp_name As String, _
                                                          ByVal r_exp_file_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=CurrentName("LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN")) = r_exp_file_full_name
End Property

'Public Property Get LastModExpFileFullName(Optional ByVal comp_name As String) As String
'    LastModExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=CurrentName("LAST_MOD_EXP_FILE_FULL_NAME"))
'End Property
'
'Public Property Let LastModExpFileFullName(Optional ByVal comp_name As String, _
'                                                          ByVal r_exp_file_full_name As String)
'    Value(pp_section:=comp_name, pp_value_name:=CurrentName("LAST_MOD_EXP_FILE_FULL_NAME")) = r_exp_file_full_name
'End Property

Public Property Let LastModWbk(Optional ByVal r_comp_name As String, _
                                        ByVal r_wbk As Workbook)
    If r_wbk Is Nothing _
    Then Value(pp_section:=r_comp_name, pp_value_name:=CurrentName("LAST_MOD_WBK_FULL_NAME")) = "unknown" _
    Else Value(pp_section:=r_comp_name, pp_value_name:=CurrentName("LAST_MOD_WBK_FULL_NAME")) = r_wbk.FullName
    
End Property

Public Property Get LastModWbkBaseName(Optional ByVal comp_name As String) As String
    Dim fso As New FileSystemObject
    LastModWbkBaseName = fso.GetBaseName(Value(pp_section:=comp_name, pp_value_name:=CurrentName("LAST_MOD_WBK_Full_NAME")))
    Set fso = Nothing
End Property

Public Property Get LastModWbkFullName(Optional ByVal comp_name As String) As String
    LastModWbkFullName = Value(pp_section:=comp_name, pp_value_name:=CurrentName("LAST_MOD_WBK_FULL_NAME"))
End Property

Public Property Get LastModWbkName(Optional ByVal comp_name As String) As String
    Dim fso As New FileSystemObject
    LastModWbkName = fso.GetFileName(Value(pp_section:=comp_name, pp_value_name:=CurrentName("LAST_MOD_WBK_FULL_NAME")))
End Property

Public Property Get RevisionNumber(Optional ByVal comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    RevisionNumber = Value(pp_section:=comp_name, pp_value_name:=CurrentName("REVISION_NUMBER"))
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
        Value(pp_section:=comp_name, pp_value_name:=CurrentName("REVISION_NUMBER")) = CompManDat.RevisionNumberInitial
    Else
        Value(pp_section:=comp_name, pp_value_name:=CurrentName("REVISION_NUMBER")) = comp_rev_no
    End If
    
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get LastModExpFile(Optional ByVal comp_name) As File
    Const PROC = "LastModExpFile Get"
    
    On Error GoTo eh
    Dim FileName    As String
    Dim sPath       As String
    
    sPath = wsConfig.FolderCommonComponentsPath & "\"
    With New FileSystemObject
        FileName = .GetFileName(LastModExpFileFullNameOrigin(comp_name))
        If FileName <> vbNullString Then
            If .FileExists(sPath & FileName) Then
                Set LastModExpFile = .GetFile(sPath & FileName)
            Else
                Set LastModExpFile = .CreateTextFile(sPath & FileName)
            End If
        End If
    End With
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let LastModExpFile(Optional ByVal comp_name, _
                                           ByVal comp_exp_file As File)
' ---------------------------------------------------------------------------
' Copies the Raw Export File from its host Workbook location to the Common
' Components Folder from where it is used as the source for the update of
' outdated Used Common Components.
' ---------------------------------------------------------------------------
    comp_name = comp_name ' not used
    comp_exp_file.Copy Destination:=wsConfig.FolderCommonComponentsPath & "\" & comp_exp_file.Name, OverWriteFiles:=True
End Property

Public Property Get LastModExpFileFullName(Optional ByVal comp_name As String) As String
' ---------------------------------------------------------------------------
' Returns the Export File Full Name which refers to the Export File saved in
' the Common-Components folder.
' ---------------------------------------------------------------------------
    With New FileSystemObject
        LastModExpFileFullName = wsConfig.FolderCommonComponentsPath & "\" & .GetFileName(LastModExpFileFullNameOrigin(comp_name))
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

Private Function CurrentName(ByVal sNameConst As String) As String
    Dim v   As Variant
    Dim dct As Dictionary
    
    Set dct = HskpngValueNamesCurrent
    For Each v In dct
        If v = sNameConst Then
            CurrentName = Split(dct(v), ":")(0)
        End If
    Next v
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCommComps" & "." & sProc
End Function

Public Function Exists(ByVal x_vbc As VBComponent, _
                       ByVal x_exp_file_ext As String, _
              Optional ByRef x_modified_by_wbk_name As String, _
              Optional ByRef x_export_file_full_name As String, _
              Optional ByRef x_export_file As File, _
              Optional ByRef x_last_mod_rev_no As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the VBComponent's (x_vbc) Export-File exists in the
' global Common-Components-Folder.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Exists"
    
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    Dim Comp    As New clsComp
    
    mBasic.BoP ErrSrc(PROC), x_vbc.Name
    sFile = wsConfig.FolderCommonComponentsPath & "\" & x_vbc.Name & x_exp_file_ext
    If fso.FileExists(sFile) Then
        Exists = True
        x_export_file_full_name = sFile
        Set x_export_file = mCommComps.LastModExpFile(x_vbc.Name)
        x_modified_by_wbk_name = mCommComps.LastModWbkName(x_vbc.Name)
        x_last_mod_rev_no = mCommComps.RevisionNumber(x_vbc.Name)
    End If
    
xt: Set Comp = Nothing
    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC), x_vbc.Name & " = " & CStr(Exists)
    
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

Public Function HskpngValueNames() As Boolean
' ----------------------------------------------------------------------------
' Renames all value names with a syntax <new>:<old> in all sections,
' Removes all value names with a syntax :<old>
' When at least one housekeeping action had been performed the function
' returns TRUE.
' ----------------------------------------------------------------------------
    Const PROC = "HskpngValueNames"
    
    On Error GoTo eh
    Dim dctNames    As Dictionary
    Dim dctSects    As Dictionary
    Dim i           As Long
    Dim sNew        As String
    Dim sOld        As String
    Dim v           As Variant
    Dim vName       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set dctNames = HskpngValueNamesCurrent
    
    For i = 0 To dctNames.Count - 1
        vName = Split(dctNames.Items(i), ":")
        If UBound(vName) = 1 Then
            '~~> <new>:<old> or :<remove>
            If vName(0) = vbNullString Then
                '~~ Remove the name in all sections
                Set dctSects = ValueSections
                For Each v In dctSects
                    If ValueNameRemoveInAllSections(vName(1), v) Then
                        HskpngValueNames = True
                    End If
                Next v
            Else
                '~~> Rename the old name with the new name in all sections
                sNew = vName(0)
                sOld = vName(1)
                If sOld <> vbNullString And sNew <> vbNullString Then
                    If ValueNameRenameInAllSections(sOld, sNew) = True Then
                        HskpngValueNames = True
                    End If
                End If
            End If
        End If
    Next i

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function HskpngValueNamesCurrent() As Dictionary

    Static dct As Dictionary
    If dct Is Nothing Then
        Set dct = New Dictionary
        dct.Add "LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN", _
                 LAST_MOD_EXP_FILE_FULL_NAME_ORIGIN
        dct.Add "LAST_MOD_WBK_FULL_NAME", _
                 LAST_MOD_WBK_FULL_NAME
        dct.Add "REVISION_NUMBER", _
                 REVISION_NUMBER
    End If
    Set HskpngValueNamesCurrent = dct
    
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
                     "It will be exported and copied to the ""Common-Components Folder"" and the 'Revision Number' will be increased."
    End With
    With Msg.Section(3)
        .Label.Text = Replace(Replace(BttnUpdate, vbLf, " "), "  ", " ")
        .Label.FontColor = rgbBlue
        .Text.Text = "The hosted version will be updated with the ""Common-Components Folder"" version (by re-import) and " & _
                     "the 'Revision-Number' of the hosted version will be set identical with the ""Common-Components Folder"" version"
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
                     "'Export File' is copied to the ""Common-Components Folder"" and the 'Revision Number' is increased " & _
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

Public Function LastModExpFileExists(ByVal comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when an Export File with the name extracted from the
' LastModExpFileFullNameOrigin exists in the Common Components Folder
' ----------------------------------------------------------------------------
    Dim FileName    As String
    With New FileSystemObject
        FileName = .GetFileName(LastModExpFileFullNameOrigin(comp_name))
        LastModExpFileExists = .FileExists(wsConfig.FolderCommonComponentsPath & "\" & FileName)
    End With
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
' and "skip forever".
' The service considers used and hosted Common Components to be updated by
' dedicated buttons and section texts.
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdateChoice"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim AppRunArgs      As Dictionary
    Dim cllButtons      As Collection
    Dim Comp            As clsComp
    Dim fUpdate         As fMsg
    Dim Msg             As mMsg.TypeMsg
    Dim sUpdate         As String
    Dim sSkipForever    As String
    Dim sAbout          As String
    Dim sSkipNow        As String
    Dim sModWbkName     As String
    Dim sUpdateBttnTxt  As String
    Dim sSkipNowNote    As String
    
    Set AppRunArgs = New Dictionary
    Qoutdated.First Comp ' Get the next outdated component from the queue
    
    With Comp
        sModWbkName = mCommComps.LastModWbkName(.CompName)
        BttnDsplyDiffs = "Display the code modifications"
        Select Case .KindOfComp
            Case enCommCompHosted
                UpdateDialogTitle = "Hosted ""Common Component"" apparently modified within a Worbook using it!"
                sUpdate = "The ""Common Component""   " & mBasic.Spaced(.CompName) & "   hosted in this " & _
                          "Workbook has been modified within the Workbook/VB-Project   " & _
                          mBasic.Spaced(sModWbkName) & " .   It needs to be updated in this hosting Workbook " & _
                          "for consistency! The extent of the coded modification can be seen with """ & _
                          BttnDsplyDiffs & """ wich should allow an estimation wether or to which extent " & _
                          "re-testing is appropriate."
                sSkipNow = "The update will be postponed and proposed with the next Workbook open."
                sSkipNowNote = "Please note: In contrast to a ""used Common Component"" it is not possible " & _
                               "to skip this update forever. The only way to achieve this is to remove the " & _
                               "hosted indication for this Workbook. With the next Workbook open the ""Common " & _
                               "Component"" will be de-registered as hosted and registered as a used one. As a " & _
                               "consequence it will remain ""not hosted"" until another Workbooks claims " & _
                               "hosting it (or forever in case none ever does)."
                BttnUpdate = "Update the hosted" & vbLf & "Common Component" & vbLf & vbLf & .CompName
                sUpdateBttnTxt = "With this update the hosting Workbook again becomes the Workbook hosting the raw " & _
                                 "version of the Common Component."
                BttnSkipForNow = "Skip this update for now"
                Set cllButtons = mMsg.Buttons(BttnUpdate, vbLf, BttnDsplyDiffs, vbLf, BttnSkipForNow)
                Set fUpdate = mMsg.MsgInstance(UpdateDialogTitle)
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

            Case enCommCompUsed
                UpdateDialogTitle = "A used ""Common Component"" is outdated, i.e. has been modified."
                sUpdate = "The used ""Common Component""   " & mBasic.Spaced(.CompName) & "   has been modified within " & _
                          "the Workbook/VB-Project (" & "" & "). Updating the used version is thus recommended!"
                sSkipNow = "Update will be postponed and proposed with the next Workbook open"
                sSkipForever = "The component, known as a potential 'Used Common Component' will be de-registered " & _
                               "and ignored in the future! I.e. the ""used"" status will be changed into ""private"" status. " & _
                               "Re-instantiating as a ""used Common Component"" will requires the following steps:" & vbLf & _
                               "1. Remove it" & vbLf & _
                               "2. Save the Workbook" & vbLf & _
                               "3. Re-Import it from the ""Common-Components"" folder."
                BttnUpdate = "Update the used" & vbLf & "Common Component" & vbLf & vbLf & .CompName
                sUpdateBttnTxt = "The outdated ""Common Component"" used in this Workbook becomes up-to-date again."
                BttnSkipForNow = "Skip this update" & vbLf & "for now"
                BttnSkipForever = "Skip this update" & vbLf & "f o r e v e r" & vbLf & "(I am aware of the consequence)"
                Set fUpdate = mMsg.MsgInstance(UpdateDialogTitle)
                Set cllButtons = mMsg.Buttons(BttnUpdate, vbLf, BttnDsplyDiffs, vbLf, BttnSkipForNow, BttnSkipForever)
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
            Case Else
                Stop
        End Select
        
    End With
    
    With Msg
        .Section(1).Text.Text = sUpdate
        With .Section(2)
            With .Label
                .Text = Replace(BttnUpdate, vbLf, " ") & ":"
                .FontColor = rgbBlue
            End With
            .Text.Text = sUpdateBttnTxt
        End With
        With .Section(3)
            With .Label
                .Text = Replace(BttnDsplyDiffs, vbLf, " ") & ":"
                .FontColor = rgbBlue
            End With
            .Text.Text = "The displayed code modifications (return with Esc) may help estimating the extent to which re-testing is appropriate after an update."
        End With
        
        With .Section(4)
            With .Label
                .Text = Replace(BttnSkipForNow, vbLf, " ") & ":"
                .FontColor = rgbBlue
            End With
            .Text.Text = sSkipNow
        End With
        If Comp.KindOfComp = enCommCompHosted Then
            With .Section(5).Text
                .Text = sSkipNowNote
                .FontBold = True
            End With
        Else
            '~~ Skip forever is only an option for Workbooks using the Common Component
            With .Section(6)
                With .Label
                    .Text = Replace(BttnSkipForever, vbLf, " ")
                    .FontColor = rgbBlue
                End With
                .Text.Text = sSkipForever
            End With
        End If
        With .Section(7)
            With .Label
                .Text = "About:"
                .OpenWhenClicked = GITHUB_REPO_URL & "#common-components"
                .FontColor = rgbBlue
                .FontUnderline = True
            End With
            .Text.Text = "A ""Common Component"" is one of which an Export-File resides in the ""Common-Components folder"". " & _
                         "It may be modified within whichever Workbook's VB-Project using it, preferrably in a Workbook " & _
                         "which claims hosting it. When a ""Common Component"" is modified its ""Revision Number"" is " & _
                         "increased and the Export-File is copied/overwritten in the ""Common-Components Folder""."
        End With
    End With
    
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
        Services.ExpFilesDiffDisplay .ExpFileFullName, mCommComps.LastModExpFile(.CompName), "Currently used (" & .ExpFileFullName & ")", "Up-to-date (" & mCommComps.LastModExpFile(.CompName).Path & ")"
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
    CompManDat.RegistrationState(u_comp_name) = enRegStatePrivate
    Qoutdated.DeQueue
    Set wbk = Services.Serviced
    With New clsComp
        .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
    End With
    
    With Services
        .NoOfItemsIgnored = .NoOfItemsIgnored + 1
        .LogServicedEntry "Outdated used Commpon Component: Update skipped forever!"
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
        .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
    End With
    
    With Services
        .NoOfItemsIgnored = .NoOfItemsIgnored + 1
        .LogServicedEntry "Outdated used Commpon Component: Update skipped for now!"
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
        .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
        v = .KindOfComp
        sFile = mCommComps.LastModExpFileFullName(.CompName)
        mUpdate.ByReImport b_wbk_target:=wbk _
                         , b_vbc_name:=u_comp_name _
                         , b_exp_file:=sFile
        
        With Services
            .NoOfItemsServiced = .NoOfItemsServiced + 1
            .NoOfItemsServicedNames = u_comp_name
            .DsplyProgress "used Common Components updated"
        End With
    
        Select Case .KindOfComp
            Case enCommCompHosted
                Services.LogServicedEntry "Outdated Common Component hosted updated by re-import of the Export-File in the Common-Components folder"
                '~~ When a hosted Common Component is updated it "again" becomes the raw host!"
                mCommComps.LastModWbk(.CompName) = .Wrkbk
                mCommComps.LastModExpFileFullNameOrigin(.CompName) = .ExpFileFullName
            Case enCommCompUsed
                Services.LogServicedEntry "Outdated Common Component used updated by re-import of the Export-File in the Common-Components folder"
        End Select
    
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
            If (.KindOfComp = enCommCompHosted Or .KindOfComp = enCommCompUsed) Then
                If .Outdated Then
                    Qoutdated.EnQueue Comp
                    sName = .CompName
                    With Services
                        .NoOfItemsServicedNames = sName
                        .NoOfItemsOutdated = Qoutdated.Size
                        .DsplyProgress "collected Common Components outdated"
                    End With
                Else
                    '~~ When not outdated due to a code difference the revision numbers ought to be equal
                    If .RevisionNumber <> mCommComps.RevisionNumber(.CompName) Then
                        .RevisionNumber = mCommComps.RevisionNumber(.CompName)
                    End If
                    With Services
                        .NoOfItemsIgnored = .NoOfItemsIgnored + 1
                        .LogServicedEntry "Used Common Component is up-to-date"
                    End With
                End If ' .Outdated
            End If
        End With
        Set Comp = Nothing
        Services.DsplyProgress "collected Common Components outdated"
    Next v
    Services.DsplyProgress "collected Common Components outdated"
    
xt: If wsService.CommonComponentsUsed = 0 Then wsService.CommonComponentsUsed = Services.NoOfCommonComponents
    If wsService.CommonComponentsOutdated = 0 Then wsService.CommonComponentsOutdated = Qoutdated.Size
    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

'Public Sub Register(ByVal r_comp_name As String, _
'                    ByVal r_exp_file As String)
'    Const PROC  As String = "Register"
'
'    Dim fso     As New FileSystemObject
'    Dim Comp    As clsComp
'
'    mBasic.BoP ErrSrc(PROC)
'    Set Comp = New clsComp
'    With Comp
'        .Wrkbk = Services.Serviced
'        .CompName = r_comp_name
'        If Not .Outdated Then
'            LastModWbk(r_comp_name) = .Wrkbk
'            LastModExpFileFullNameOrigin(r_comp_name) = r_exp_file
'        End If
'    End With
'    Set fso = Nothing
'
'xt: mBasic.EoP ErrSrc(PROC)
'End Sub

Public Sub SaveToCommonComponentsFolder(ByVal s_comp_name As String, _
                                        ByVal s_file As File, _
                                        ByVal s_file_full_name As String, _
                                        ByVal s_raw_host As Workbook)
' ------------------------------------------------------------------------------
' Save a copy of the hosted raw`s (s_comp_name) export file to the Common
' Components folder which serves as source for the update of Common Components
' used in other VB-Projects.
' ------------------------------------------------------------------------------
    Const PROC  As String = "SaveToCommonComponentsFolder"
    
    Dim frxFile As File
    Dim fso     As New FileSystemObject
    
    mBasic.BoP ErrSrc(PROC)
    mCommComps.LastModExpFile(s_comp_name) = s_file
    '~~ When the Export file has a .frm extension the .frx file needs to be copied too
    If fso.GetExtensionName(s_file_full_name) = "frm" Then
        Set frxFile = fso.GetFile(Replace(s_file_full_name, "frm", "frx"))
        mCommComps.LastModExpFile(s_comp_name) = frxFile
    End If
    
xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)

End Sub

Private Function ValueNameRemoveInAllSections(ByVal pp_name As String, _
                                     Optional ByVal pp_section As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim dct As Dictionary
    Dim v   As Variant
    
    Set dct = ValueSections
    
    For Each v In dct
        If mFso.PPremoveNames(pp_file:=CommCompsDatFileFullName _
                            , pp_section:=v _
                            , pp_value_names:=pp_name) Then
            ValueNameRemoveInAllSections = True
        End If
    Next v
    
End Function

Private Function ValueNameRenameInAllSections(ByVal v_old As String, _
                                              ByVal v_new As String) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim dct     As Dictionary
    Dim v       As Variant
    Dim bReorg  As Boolean
    Dim i       As Long
    
    Set dct = ValueSections
    For Each v In dct
        i = i + 1
        bReorg = i = dct.Count
        If mFso.PPvalueNameRename(v_old, v_new, CommCompsDatFileFullName, v, bReorg) Then
            ValueNameRenameInAllSections = True
        End If
    Next v
    
End Function

Private Function ValueSections() As Dictionary
    Static dct As Dictionary
    If dct Is Nothing Then Set dct = mFso.PPsections(CommCompsDatFileFullName)
    Set ValueSections = dct
End Function

