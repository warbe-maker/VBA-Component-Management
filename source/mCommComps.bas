Attribute VB_Name = "mCommComps"
Option Explicit

Public Qoutdated                                As clsQ

Private BttnUpdate                              As String
Private BttnDsplyDiffs                          As String
Private BttnSkipForNow                          As String
Private BttnSkipForever                         As String
Private UpdateDialogTitle                       As String
Private UpdateDialogTop                         As Long
Private UpdateDialogLeft                        As Long

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mCommComps" & "." & es_proc
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
    Dim AppRunArgs      As Dictionary
    Dim cllButtons      As Collection
    Dim Comp            As clsComp
    Dim fUpdate         As fMsg
    Dim Msg             As mMsg.TypeMsg
    Dim sUpdate         As String
    Dim sSkipForever    As String
    Dim sSkipNow        As String
    Dim sModWbkName     As String
    Dim sUpdateBttnTxt  As String
    Dim sSkipNowNote    As String
    
    Set AppRunArgs = New Dictionary
    Qoutdated.First Comp ' Get the next outdated component from the queue
    
    With Comp
        sModWbkName = CommComps.LastModWbkName(.CompName)
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
                                            , "OutdatedUpdateChoiceUpdate" _
                                            , Comp.CompName
                mMsg.BttnAppRun AppRunArgs, BttnDsplyDiffs _
                                            , ThisWorkbook _
                                            , "OutdatedUpdateChoiceDsplyDiffs"
                mMsg.BttnAppRun AppRunArgs, BttnSkipForNow _
                                            , ThisWorkbook _
                                            , "OutdatedUpdateChoiceSkipForNow" _
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
                                            , "OutdatedUpdateChoiceUpdate" _
                                            , Comp.CompName
                mMsg.BttnAppRun AppRunArgs, BttnDsplyDiffs _
                                            , ThisWorkbook _
                                            , "OutdatedUpdateChoiceDsplyDiffs"
                mMsg.BttnAppRun AppRunArgs, BttnSkipForNow _
                                            , ThisWorkbook _
                                            , "OutdatedUpdateChoiceSkipForNow" _
                                            , Comp.CompName
                mMsg.BttnAppRun AppRunArgs, BttnSkipForever _
                                            , ThisWorkbook _
                                            , "OutdatedUpdateChoiceSkipForever" _
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
        Services.ExpFilesDiffDisplay .ExpFileFullName, CommComps.LastModExpFile(.CompName), "Currently used (" & .ExpFileFullName & ")", "Up-to-date (" & CommComps.LastModExpFile(.CompName).Path & ")"
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
    OutdatedUpdate
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
    OutdatedUpdate
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
        sFile = CommComps.LastModExpFileFullName(.CompName)
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
                CommComps.LastModWbk(.CompName) = .Wrkbk
                CommComps.LastModExpFileFullNameOrigin(.CompName) = .ExpFileFullName
            Case enCommCompUsed
                Services.LogServicedEntry "Outdated Common Component used updated by re-import of the Export-File in the Common-Components folder"
        End Select
    
    End With
    Qoutdated.DeQueue
    Set Comp = Nothing
    
xt: Services.MessageUnload UpdateDialogTitle
    OutdatedUpdate
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
                    If .RevisionNumber <> CommComps.RevisionNumber(.CompName) Then
                        .RevisionNumber = CommComps.RevisionNumber(.CompName)
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
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub



