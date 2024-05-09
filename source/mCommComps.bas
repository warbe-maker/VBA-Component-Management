Attribute VB_Name = "mCommComps"
Option Explicit
' ------------------------------------------------------------------------
' Standard-Module mCommComps: Services for Common Components.
' ======================
'
' Public services:
' ----------------
' OutdatedUpdate Collects all used outdated Common Components when called
'                for the first time and displays the first one queued in
'                Qoutdated. The service is re-called until the queue is
'                empty. The display of the update choices is a mode-less
'                dialog which calls sub-services in accordance with the
'                button pressed.
'
' W. Rauschenberger, Berlin Jul 18 2023
' ------------------------------------------------------------------------
Public Qoutdated                                As clsQ

Private BttnUpdate                              As String
Private BttnDsplyDiffs                          As String
Private BttnSkipForNow                          As String
Private BttnSkipForever                         As String
Private UpdateTitle                             As String
Private UpdateDialogTop                         As Long
Private UpdateDialogLeft                        As Long
Private sUpdateDone                             As String

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
    
    If Qoutdated Is Nothing Then
        OutdatedUpdate1Collect
                
        If Not Qoutdated.IsEmpty Then sUpdateDone = "updated" Else sUpdateDone = "outdated"
    End If
    If Not Qoutdated.IsEmpty Then
        OutdatedUpdate2Choice
    Else
        Services.LogEntrySummary Application.StatusBar
    End If
    
xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdate2Choice()
' ------------------------------------------------------------------------------
' Displays the first outdated Common Component in the queue Qoutdated in a mode-
' less dialog for one of the options: "update", "display diffs", "skip for now",
' and "skip forever".
' The service considers used and hosted Common Components to be updated by
' dedicated buttons and section texts.
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdate2Choice"
    
    On Error GoTo eh
    Dim AppRunArgs          As Dictionary
    Dim cllButtons          As Collection
    Dim Comp                As clsComp
    Dim fUpdate             As fMsg
    Dim i                   As Long
    Dim sModWbkName         As String
    Dim sSkipForeverNote    As String
    Dim sSkipForNow         As String
    Dim sSkipForNowNote     As String
    Dim sUpdateNote         As String
    Dim sUpdateBttnTxt      As String
    
    BttnDsplyDiffs = "Display the code modifications"
    BttnSkipForNow = "Skip this update" & vbLf & "for now"
    BttnSkipForever = "Skip this update" & vbLf & mBasic.Spaced("forever")
    
    Set AppRunArgs = New Dictionary
    Qoutdated.First Comp ' Get the next outdated component from the queue
    
    With Comp
        sModWbkName = CommComps.LastModInWbkName(.CompName)
        Select Case .KindOfComp
            Case enCommCompHosted
                UpdateTitle = "Hosted Common Component has been modified within a Worbook/VB-Project using it!"
                sUpdateNote = "The Common Component   " & mBasic.Spaced(.CompName) & "   hosted in this " & _
                              "Workbook has been modified by/within the Workbook/VB-Project   " & _
                              mBasic.Spaced(sModWbkName) & " .   The extent of the modification may be " & _
                              "displayed in order to allow an estimation wether or not (or to which extent) " & _
                              "re-testing or performing a regression test is appropriate - suposed the ""hosting"" " & _
                              "Workbook provides such a test environment."
                sSkipForNow = "The update will be postponed and proposed with the next Workbook open."
                sSkipForNowNote = "Please note: In contrast to a ""used"" Common Component it is not possible " & _
                                  "to skip this update forever. The only way to achieve this is to remove the " & _
                                  """hosted"" indication for this Workbook. Subsquently with the next Workbook " & _
                                  "open the (then) ""used"" Common Component's update may be skipped forever, " & _
                                  "i.e. it will be de-registered as ""used"" and registered as ""private"". It " & _
                                  "also should be noted that the Common Component will remain ""not hosted"" until " & _
                                  "another Workbooks claims ""hosting"" it."
                BttnUpdate = "Update the hosted" & vbLf & "Common Component" & vbLf & vbLf & .CompName
                sUpdateBttnTxt = "With this update the ""hosted"" Common Component will become in sync with all Workbook using this version."
                Set cllButtons = mMsg.Buttons(BttnUpdate, vbLf, BttnDsplyDiffs, vbLf, BttnSkipForNow)

            Case enCommCompUsed
                UpdateTitle = "Used Common Component is outdated!"
                sUpdateNote = "The used Common Component   " & mBasic.Spaced(.CompName) & "   has been modified within " & _
                          "the Workbook/VB-Project   " & mBasic.Spaced(sModWbkName) & " .  Updating the used version " & _
                          "is thus recommended!"
                sSkipForNow = "Update will be postponed and proposed with the next Workbook open"
                sSkipForeverNote = "The component, currently known as a used Common Component will be de-registered " & _
                               "and ignored in the future! I.e. its ""used"" status will be changed into ""private"". " & _
                               "Re-instantiating it as a ""used"" Common Component will requires the following steps:" & vbLf & _
                               "1. Remove it" & vbLf & _
                               "2. Save the Workbook" & vbLf & _
                               "3. Re-Import it from the ""Common-Components"" folder."
                BttnUpdate = "Update the used" & vbLf & "Common Component" & vbLf & vbLf & .CompName
                sUpdateBttnTxt = "The outdated ""used"" Common Component becomes up-to-date again."
                Set cllButtons = mMsg.Buttons(BttnUpdate, vbLf, BttnDsplyDiffs, vbLf, BttnSkipForNow, BttnSkipForever)
            Case Else
                Stop
        End Select
        
        Set fUpdate = mMsg.MsgInstance(UpdateTitle)
         '~~ AppRun arguments for potentially displayed buttons
        mMsg.BttnAppRun AppRunArgs, BttnUpdate _
                                    , ThisWorkbook _
                                    , "OutdatedUpdate2Choice1Update" _
                                    , .CompName
        mMsg.BttnAppRun AppRunArgs, BttnDsplyDiffs _
                                    , ThisWorkbook _
                                    , "OutdatedUpdate2Choice2DsplyDiffs"
        mMsg.BttnAppRun AppRunArgs, BttnSkipForNow _
                                    , ThisWorkbook _
                                    , "OutdatedUpdate2Choice3SkipForNow" _
                                    , .CompName
        mMsg.BttnAppRun AppRunArgs, BttnSkipForever _
                                    , ThisWorkbook _
                                    , "OutdatedUpdate2Choice4SkipForever" _
                                    , .CompName
    End With
    
    mCompMan.MsgInit
    i = 0
    With Msg
        i = i + 1:  .Section(i).Text.Text = sUpdateNote
        i = i + 1:  With .Section(i)
                        .Label.Text = "Update:"
                        .Text.Text = sUpdateBttnTxt
                    End With
        i = i + 1:  With .Section(i)
                        .Label.Text = "Display difference:"
                        .Text.Text = "The displayed code modifications (return with Esc) may help estimating the extent to which re-testing is appropriate after an update."
                    End With
        i = i + 1:  With .Section(i)
                        .Label.Text = "Skip for now:"
                        .Text.Text = sSkipForNow
                    End With
        i = i + 1:  With .Section(i).Text
                        .Text = sSkipForNowNote
                        .FontBold = True
                    End With
        If Comp.KindOfComp = enCommCompUsed Then
            '~~ Skip forever is an option only for used Common Components (hosted need to be de-registered as such first)
            i = i + 1:  With .Section(i)
                            .Label.Text = "Skip forever:"
                            .Text.Text = sSkipForeverNote
                        End With
        End If
        i = i + 1:  With .Section(i).Label
                        .Text = "About:"
                        .OnClickAction = GITHUB_REPO_URL & "#common-components"
                        .FontUnderline = True
                    End With
                    .Section(i).Text.Text = _
                        "A 'Common Component' is one of which the Export-File is copied to the 'Common-Components folder' when the code is modified, " & _
                        "whereby it may be modified within whichever Workbook's VB-Project using it, preferrably in a Workbook " & _
                        "which claims 'hosting' it. When a 'Common Component' is modified its 'Revision Number' is " & _
                        "increased and the Export-File is copied/overwritten in the 'Common-Components Folder'."
     End With
    
    mMsg.Dsply dsply_title:=UpdateTitle _
                 , dsply_msg:=Msg _
                 , dsply_Label_spec:="R70" _
                 , dsply_buttons:=cllButtons _
                 , dsply_modeless:=True _
                 , dsply_buttons_app_run:=AppRunArgs _
                 , dsply_width_min:=45 _
                 , dsply_pos:=UpdateDialogTop & ";" & UpdateDialogLeft
    DoEvents
    
xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdate2Choice2DsplyDiffs()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdate2Choice2DsplyDiffs"
    
    On Error GoTo eh
    Dim Comp    As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.First Comp
    With CommComps
        .CompName = Comp.CompName
        .DsplyDiffsOfTheCurrentModificationsVersusTheCurrentPubliicCode Comp
    End With
    Set Comp = Nothing
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdate2Choice4SkipForever(ByVal u_comp_name)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdate2Choice4SkipForever"
    
    On Error GoTo eh
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    PPCompManDat.RegistrationState(u_comp_name) = enRegStatePrivate
    Qoutdated.DeQueue
    Set wbk = Services.ServicedWbk
    With New clsComp
        .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
    End With
    
    With Services
        .NoOfItemsSkipped = .NoOfItemsSkipped + 1
        .ServicedItemLogEntry "Outdated used Commpon Component: Update skipped forever!"
    End With
    
xt: Services.MessageUnload UpdateTitle
    OutdatedUpdate
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdate2Choice3SkipForNow(ByVal u_comp_name As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdate2Choice3SkipForNow"
    
    On Error GoTo eh
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.DeQueue
    Set wbk = Services.ServicedWbk
    With New clsComp
        .Wrkbk = wbk
        .CompName = u_comp_name
        Services.ServicedItem = .VBComp
    End With
    
    With Services
        .NoOfItemsSkipped = .NoOfItemsSkipped + 1
        .ServicedItemLogEntry "Outdated used Commpon Component: Update skipped for now!"
    End With
    
xt: Services.MessageUnload UpdateTitle
    OutdatedUpdate
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdate2Choice1Update(ByVal u_comp_name As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdate2Choice1Update"
    
    On Error GoTo eh
    Dim wbk     As Workbook
    Dim Comp    As clsComp
    Dim v       As Variant
    Dim sFile   As String
    
    mBasic.BoP ErrSrc(PROC)
    Qoutdated.First Comp
    Set wbk = Services.ServicedWbk
    Set Comp = New clsComp
    With Comp
        .Wrkbk = wbk
        .CompName = u_comp_name
        CommComps.CompName = u_comp_name
        Services.ServicedItem = .VBComp
        v = .KindOfComp
        sFile = CommComps.LastModExpFileFullName
        mUpdate.ByReImport b_wbk_target:=wbk _
                         , b_vbc_name:=u_comp_name _
                         , b_exp_file:=sFile
        
        With Services
            .NoOfItemsServiced = .NoOfItemsServiced + 1
            .NoOfItemsServicedNames = u_comp_name
            .Progress "Common Components updated"
        End With
    
        Select Case .KindOfComp
            Case enCommCompHosted
                Services.ServicedItemLogEntry "Outdated hosted Common Component - last modified in   " & mBasic.Spaced(CommComps.LastModInWbkName) & "   - updated by re-import of the Export-File in the Common-Components folder"
                '~~ When a hosted Common Component is updated it "again" becomes the raw host!"
                CommComps.LastModInWbk = .Wrkbk
                CommComps.LastModExpFileFullNameOrigin = .ExpFileFullName
            Case enCommCompUsed
                Services.ServicedItemLogEntry "Outdated used Common Component - last modified in   " & mBasic.Spaced(CommComps.LastModInWbkName) & "   - updated by re-import of the Export-File in the Common-Components folder"
        End Select
    
    End With
    Qoutdated.DeQueue
    Set Comp = Nothing
    
xt: Services.MessageUnload UpdateTitle
    OutdatedUpdate
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub OutdatedUpdate1Collect()
' ------------------------------------------------------------------------------
' Collects all outdated Used Common Components and enqueues them in Qoutdated.
' ------------------------------------------------------------------------------
    Const PROC = "OutdatedUpdate1Collect"
    
    On Error GoTo eh
    Dim Comp            As clsComp
    Dim wbk             As Workbook
    Dim dct             As Dictionary
    Dim v               As Variant
    Dim sName           As String
    Dim bOutdated       As Boolean
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.ServicedWbk
    Set Qoutdated = New clsQ
    Set dct = Comps.All ' all = all relevant for the current service
    Prgrss.ItemsTotal = dct.Count
    Prgrss.Operation = "outdated"
    
    For Each v In dct
        Set Comp = dct(v)
        With Comp
            CommComps.CompName = .CompName
            Services.ServicedItem = .VBComp
            If (.KindOfComp = enCommCompHosted Or .KindOfComp = enCommCompUsed) Then
                bOutdated = .Outdated
                If bOutdated Then
                    Qoutdated.EnQueue Comp
                    sName = .CompName
                    Prgrss.ItemDone = .CompName
                    With Services
                        .NoOfItemsServicedNames = sName
                        .NoOfItemsOutdated = Qoutdated.Size
                    End With
                Else
                    Prgrss.ItemSkipped
                    If .LastModAtDateTimeUTC < CommComps.LastModAtDateTimeUTC Then
                        .LastModAtDateTimeUTC = CommComps.LastModAtDateTimeUTC
                    End If
                    With Services
                        .NoOfItemsSkipped = .NoOfItemsSkipped + 1
                        .ServicedItemLogEntry "Used Common Component is up-to-date"
                    End With
                End If ' .Outdated
            Else
                Prgrss.ItemSkipped
            End If
        End With
        Set Comp = Nothing
        Application.StatusBar = vbNullString
        DoEvents
    Next v
    mCompManMenu.Setup
    Prgrss.Dsply
    
xt: If wsService.CommonComponentsUsed = 0 Then wsService.CommonComponentsUsed = Services.NoOfCommonComponents
    If wsService.CommonComponentsOutdated = 0 Then wsService.CommonComponentsOutdated = Qoutdated.Size
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

