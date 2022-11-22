Attribute VB_Name = "mComCompRawsHosted"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mComCompRawsHosted
' Maintains in a file named ComCompsHosted.dat for each Workbook which hosts
' at least one Raw Common Component with the following structure:
'
' [<component-name]
' ExpFileFullName=<export-file-full-name>
' RevisionNumber=yyyy-mm-dd.nnn
'
' The entries (sections) are maintained along with the Workbook_BeforeSave
' event via the ExportChangedComponents service. The revision number is
' increased whith each saved modification.
' ---------------------------------------------------------------------------
Private Const VNAME_RAW_REVISION_NUMBER     As String = "RawRevisionNumber"
Private Const VNAME_RAW_EXP_FILE_FULL_NAME  As String = "RawExpFileFullName"


Public Sub Manage(ByVal mh_hosted As String)
' ----------------------------------------------------------------------------
' Manages Hosted Raw Common Components by:
' - Registering the Workbook as 'Raw-Host' when it hosts at least one Common
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
    Const PROC = "Manage"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim fso             As New FileSystemObject
    Dim RawComComp      As clsComp
    Dim sHostBaseName   As String
    
    mBasic.BoP ErrSrc(PROC)
    sHostBaseName = fso.GetBaseName(mService.WbkServiced.FullName)
    Set dctHostedRaws = New Dictionary
    HostedRaws = mh_hosted
                    
    If HostedRaws.Count <> 0 Then
        For Each v In HostedRaws
            Set RawComComp = New clsComp
            With RawComComp
                Set .Wrkbk = mService.WbkServiced
                .CompName = v
                If mComCompRawsHosted.IsNewHostedRaw(v) Then
                    '~~ Initially export the component the serviced Workbook claims 'hosted raw common component'
                    If Not fso.FileExists(.ExpFileFullName) Then
                        .VBComp.Export .ExpFileFullName
                    End If
                    mComCompRawsHosted.RawExpFileFullName(v) = .ExpFileFullName ' in any case update the Export File name
                    mComCompRawsHosted.RawRevisionNumberIncrease v     ' this will initially set it
                    mComCompRawsHosted.SaveToGlobalFolder v, .ExpFile, .ExpFileFullName
                End If
                mComCompRawsHosted.RawExpFileFullName(v) = .ExpFileFullName ' in any case update the Export File name
                If Not mComCompRawsGlobal.Exists(v) Then
                    '~~ Initially registers the Common Component in the ComComps-RawsSaved.dat file
                    '~~ Note: The Raw's Revision Number is updated whenever the raw is exported because it had been modified
                    mComCompRawsGlobal.RawHostWbFullName(v) = mService.WbkServiced.FullName
                    mComCompRawsGlobal.RawHostWbName(v) = mService.WbkServiced.name
                    mComCompRawsGlobal.RawHostWbBaseName(v) = fso.GetBaseName(mService.WbkServiced.FullName)
                    Log.Entry = "Raw-Component '" & v & "' hosted in this Workbook registered"
                End If
                If Not mComCompRawsGlobal.Exists(v) Then
                    '~~ Initially registers the Common Component in the ComComps-RawsSaved.dat file
                    '~~ Note: The Raw's Revision Number is updated whenever the raw is exported because it had been modified
                    mComCompRawsGlobal.RawHostWbFullName(v) = mService.WbkServiced.FullName
                    mComCompRawsGlobal.RawHostWbName(v) = mService.WbkServiced.name
                    mComCompRawsGlobal.RawHostWbBaseName(v) = fso.GetBaseName(mService.WbkServiced.FullName)
                    Log.Entry = "Raw-Component '" & v & "' hosted in this Workbook registered"
                ElseIf StrComp(mComCompRawsGlobal.RawHostWbFullName(v), mService.WbkServiced.FullName, vbTextCompare) <> 0 _
                    Or StrComp(mComCompRawsGlobal.RawHostWbName(v), mService.WbkServiced.name, vbTextCompare) <> 0 Then
                    '~~ Update the properties when they had changed - which may happen when the Raw Common Component's
                    '~~ host has changed
                    '~~ Note: The RevisionNumber is updated whenever the modified raw is exported
                    mComCompRawsHosted.SaveToGlobalFolder v, .ExpFile, .ExpFileFullName
                    mComCompRawsGlobal.RawHostWbFullName(v) = mService.WbkServiced.FullName
                    mComCompRawsGlobal.RawHostWbName(v) = mService.WbkServiced.name
                    mComCompRawsGlobal.RawHostWbBaseName(v) = fso.GetBaseName(mService.WbkServiced.FullName)
                    Log.Entry = "Raw Common Component '" & v & "' hosted changed properties updated"
                End If
                If mService.FilesDiffer(fd_exp_file_1:=.ExpFile _
                                      , fd_exp_file_2:=mComCompRawsGlobal.SavedExpFile(v)) Then
                    '~~ Attention! This is a cruical issue which shold never be the case. However, when different
                    '~~ computers/users are involved in the development process ...
                    '~~ Instead of simply updating the saved raw Export File better have carefully checked the case
                    If mComCompRawsGlobal.RawSavedRevisionNumber(v) = mComCompRawsHosted.RawRevisionNumber(v) Then
                        If InconsitencyWarning _
                           (sri_raw_exp_file_full_name:=.ExpFile.Path _
                          , sri_saved_exp_file_full_name:=mComCompRawsGlobal.SavedExpFile(v).Path _
                          , sri_diff_message:="While the Revision Number of the 'Hosted Raw'  " & mBasic.Spaced(v) & "  is identical with the " & _
                                              "'Saved Raw' their Export Files are different. Compared were:" & vbLf & _
                                              "Hosted Raw Export File = " & .ExpFile.Path & vbLf & _
                                              "Saved Raw Export File  = " & mComCompRawsGlobal.SavedExpFile(v).Path & vbLf & _
                                              "whereby any empty code lines and case differences had been ignored. " & _
                                              "The difference thus really matters!" _
                           ) Then
                            mComCompRawsHosted.SaveToGlobalFolder v, .ExpFile, .ExpFileFullName
                        End If
                    ElseIf mComCompRawsGlobal.RawSavedRevisionNumber(v) <> mComCompRawsHosted.RawRevisionNumber(v) Then
                        If InconsitencyWarning _
                           (sri_raw_exp_file_full_name:=.ExpFile.Path _
                          , sri_saved_exp_file_full_name:=mComCompRawsGlobal.SavedExpFile(v).Path _
                          , sri_diff_message:="The 'Revision Number' of the 'Hosted Raw Common Component's Export File' and the " & _
                                              "the 'Saved Raw's Export File' differ:" & vbLf & _
                                              "Hosted Raw = " & mComCompRawsHosted.RawRevisionNumber(v) & vbLf & _
                                              "Saved Raw  = " & mComCompRawsGlobal.RawSavedRevisionNumber(v) & vbLf & _
                                              "and also the Export Files differ. Compared were:" & vbLf & _
                                              "Hosted Raw = " & .ExpFile.Path & vbLf & _
                                              "Saved Raw  = " & mComCompRawsGlobal.SavedExpFile(v).Path & vbLf & _
                                              "whereby any empty code lines and case differences had been ignored. " & _
                                              "The difference thus really matters! Updating is not at all " & _
                                              "recommendable before the issue had been clarified." _
                           ) Then
                            mComCompRawsHosted.SaveToGlobalFolder v, .ExpFile, .ExpFileFullName
                        End If
                    End If
                End If
            End With
            Set RawComComp = Nothing
        Next v
    Else
        '~~ When this Workbook not or no longer hosts any Common Component Raws the correponding entries
        '~~ the ComCompsHosted.dat are deleted
        mFso.FileDelete mComCompRawsHosted.ComCompsHostedFileFullName
        '~~ The destiny of the corresponding data in the ComComps-Saved.dat is un-clear
        '~~ The component may be now hosted in another Workbook (likely) or the life of the
        '~~ Common Component has ended. The entry will be removed when it still points to this
        '~~ Workbook. When it points to another one it appears to have been moved alrerady.
        For Each v In mComCompRawsGlobal.Components
            If StrComp(mComCompRawsGlobal.RawHostWbFullName(comp_name:=v), mService.WbkServiced.FullName, vbTextCompare) = 0 Then
                mComCompRawsGlobal.Remove comp_name:=v
                Log.Entry = "Component no longer hosted in '" & mService.WbkServiced.FullName & "' removed from '" & mComCompRawsGlobal.ComCompsSavedFileFullName & "'"
            End If
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

Private Function InconsitencyWarning(ByVal sri_raw_exp_file_full_name, _
                                     ByVal sri_saved_exp_file_full_name, _
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
                mService.ExpFilesDiffDisplay fd_exp_file_left_full_name:=sri_raw_exp_file_full_name _
                                               , fd_exp_file_left_title:="Raw Common Component's Export File: (" & sri_raw_exp_file_full_name & ")" _
                                               , fd_exp_file_right_full_name:=sri_saved_exp_file_full_name _
                                               , fd_exp_file_right_title:="Saved Raw's Export File (" & sri_saved_exp_file_full_name & ")"
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

Private Function IsNewHostedRaw(ByVal in_vbc_name As String) As Boolean
' ---------------------------------------------------------------------------
' Returns TRUE when the serviced Workbook claims a component "hosted raw
' common coponent" and the concerned component is yet not registered in the
' Serviced Workbook's 'ComCompsHosted.dat' (ComCompsHostedFileFullName) or no
' 'ComCompsHosted.dat' exists.
' ---------------------------------------------------------------------------
    IsNewHostedRaw = Not IsRegisteredLocally(in_vbc_name)
End Function

Private Property Get IsRegisteredLocally(ByVal irl_vbc_name As String) As Boolean
    IsRegisteredLocally = Exists(irl_vbc_name)
End Property

Public Sub SaveToGlobalFolder(ByVal stgf_vbc_name As String, _
                              ByVal stgf_exp_file As File, _
                              ByVal stgf_exp_file_full_name As String)
' ------------------------------------------------------------------------------
' Save a copy of the hosted raw`s (stgf_vbc_name) export file to the Common
' Components folder which serves as source for the update of Common Components
' used in other VB-Projects.
' ------------------------------------------------------------------------------
    Dim frxFile As File
    Dim fso     As New FileSystemObject
    
    mComCompRawsGlobal.SavedExpFile(stgf_vbc_name) = stgf_exp_file
    '~~ When the Export file has a .frm extension the .frx file needs to be copied too
    If fso.GetExtensionName(stgf_exp_file_full_name) = "frm" Then
        Set frxFile = fso.GetFile(Replace(stgf_exp_file_full_name, "frm", "frx"))
        mComCompRawsGlobal.SavedExpFile(stgf_vbc_name) = frxFile
    End If

    mComCompRawsGlobal.RawExpFileFullName(stgf_vbc_name) = stgf_exp_file_full_name
    mComCompRawsGlobal.RawHostWbBaseName(stgf_vbc_name) = fso.GetBaseName(mService.WbkServiced.FullName)
    mComCompRawsGlobal.RawHostWbFullName(stgf_vbc_name) = mService.WbkServiced.FullName
    mComCompRawsGlobal.RawHostWbName(stgf_vbc_name) = mService.WbkServiced.name
    mComCompRawsGlobal.RawSavedRevisionNumber(stgf_vbc_name) = mComCompRawsHosted.RawRevisionNumber(stgf_vbc_name)
    
    Set fso = Nothing
End Sub

Private Property Get IsRegisteredGlobally(ByVal irg_vbc_name As String) As Boolean
    IsRegisteredGlobally = mComCompRawsGlobal.Exists(irg_vbc_name)
End Property

Private Property Get ComCompsHostedFileFullName() As String
    Dim wbk As Workbook
    Dim fso As New FileSystemObject
    
    Set wbk = mService.WbkServiced
    ComCompsHostedFileFullName = Replace(wbk.FullName, wbk.name, "ComCompsHosted.dat")
    If Not fso.FileExists(ComCompsHostedFileFullName) Then
        fso.CreateTextFile ComCompsHostedFileFullName
    End If
    Set fso = Nothing
    
End Property

Private Property Get RawExpFileFullName(Optional ByVal comp_name As String) As String
    RawExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME)
End Property

Private Property Let RawExpFileFullName(Optional ByVal comp_name As String, _
                                                ByVal exp_file_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME) = exp_file_full_name
End Property

Private Property Get RawRevisionNumber(Optional ByVal comp_name As String) As String
    RawRevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER)
End Property

Private Property Let RawRevisionNumber(Optional ByVal comp_name As String, _
                                               ByVal comp_rev_no As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER) = comp_rev_no
End Property

Private Property Get Value(Optional ByVal pp_section As String, _
                           Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file ComCompsHostedFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFso.FilePrivProfValue(pp_file:=ComCompsHostedFileFullName _
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
' ----------------------------------------------------------------------------
' Writes the value (pp_value) under the name (pp_value_name) into the
' ComCompsHostedFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFso.FilePrivProfValue(pp_file:=ComCompsHostedFileFullName _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Function Components() As Dictionary
    Set Components = mFso.FilePrivProfSectionNames(ComCompsHostedFileFullName)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRaw" & "." & sProc
End Function

Private Function Exists(ByVal raw_vbc_name As String) As Boolean
    Exists = Components.Exists(raw_vbc_name)
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

Private Sub Remove(ByVal comp_name As String)
    mFso.FilePrivProfRemoveSections pp_file:=ComCompsHostedFileFullName _
                                  , pp_sections:=comp_name
End Sub

Public Sub RawRevisionNumberIncrease(ByVal comp_name As String)
' ----------------------------------------------------------------------------
' Increases the revision number by one starting with 1 for a new day.
' ----------------------------------------------------------------------------
    Dim RevNo   As Long
    Dim RevDate As String
    
    If RawRevisionNumber(comp_name) = vbNullString Then
        RevNo = 1
    Else
        RevNo = Split(RawRevisionNumber(comp_name), ".")(1)
        RevDate = Split(RawRevisionNumber(comp_name), ".")(0)
        If RevDate <> Format(Now(), "YYYY-MM-DD") _
        Then RevNo = 1 _
        Else RevNo = RevNo + 1
    End If
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER) = Format(Now(), "YYYY-MM-DD") & "." & Format(RevNo, "000")

End Sub

