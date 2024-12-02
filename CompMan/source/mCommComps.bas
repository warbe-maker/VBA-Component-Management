Attribute VB_Name = "mCommComps"
Option Explicit
' ------------------------------------------------------------------------
' Standard-Module mCommComps: Services for Common Components.
' ======================
'
' Public services:
' ----------------
' Update Collects all used outdated Common Components and display choices
' for each collected.
'
' W. Rauschenberger, Berlin Jul 18 2023
' ------------------------------------------------------------------------
Private BttnSkipForever                         As String
Private BttnSkipForNow                          As String
Private BttnUpdate                              As String
Private Qoutdated                               As clsQ
Private sUpdateDone                             As String
Private UpdateDialogLeft                        As Long
Private UpdateDialogTop                         As Long
Private UpdateTitle                             As String

Private Function NextSect(ByRef n_sect As Long) As Long
    n_sect = n_sect + 1
    NextSect = n_sect
End Function

Private Sub ChoiceLoop()
' ------------------------------------------------------------------------------
' Loops through enqued outdate Common Component displaying a dialog with the
' choices: update, display diffs, skip for now, and skip forever.
' ------------------------------------------------------------------------------
    Const PROC = "ChoiceLoop"
    
    On Error GoTo eh
    Dim cllButtons          As Collection
    Dim Comp                As clsComp
    Dim fUpdate             As fMsg
    Dim i                   As Long
    Dim sComp               As String
    Dim sModWbkName         As String
    Dim sSkipForeverNote    As String
    Dim sSkipForNow         As String
    Dim sSkipForNowNote     As String
    Dim sUpdateNote         As String
    Dim sUpdateBttnTxt      As String
    
    Prgrss.ItemsTotal = Qoutdated.Size
    Prgrss.Operation = "outdated updated"
    BttnSkipForNow = "Skip this update" & vbLf & "for now"
    BttnSkipForever = "Skip this update" & vbLf & mBasic.Spaced("forever")
    
    Do While Not Qoutdated.IsEmpty
        Qoutdated.First Comp ' Get the next outdated component from the queue
        With Comp
            sComp = .CompName
            sModWbkName = CommonPublic.LastModInWrkbkFullName(sComp)
            Select Case CommonServiced.KindOfComponent(sComp)
                Case enCompCommonHosted
                    UpdateTitle = "Hosted Common Component is not/no longer up-to-date!"
                    sUpdateNote = "The hosted Common Component   " & mBasic.Spaced(sComp) & "   is outdated."
                    sSkipForNow = "The update will be postponed and proposed with the next Workbook open."
                    sSkipForNowNote = "Please note: When the update is postponed the ""used"" Common Component remains outdated! " & _
                                      "When it - by accident - is modified, this modification will never become public. " & _
                                      "A modified non-up-to-date Common Component will not be registered ""pending release"". " & _
                                      "With the next Workbook open an update will again be suggested - and when done any modification " & _
                                      "will be discarded."
                    BttnUpdate = "Update"
                    sUpdateBttnTxt = "With this update the ""hosted"" Common Component will become in sync with all Workbook using this version."
                    Set cllButtons = mMsg.Buttons(BttnUpdate, mDiff.ServicedExportVersusPublicBttn, vbLf, BttnSkipForNow)
    
                Case enCompCommonUsed
                    UpdateTitle = "Used ""Common Component"" is not/no longer up-to-date!"
                    sUpdateNote = "The used Common Component   " & mBasic.Spaced(sComp) & "   is outdated."
                    sSkipForNow = "Update will be postponed and proposed with the next Workbook open. While not up-to-date " & _
                                  "any code modification will be discarded with any subsequent export service."
                    sSkipForeverNote = "The component, currently known as a used ""Common Component"", will be de-registered " & _
                                       "and no longer updated in the future (its ""used"" status will be changed into ""private""). " & _
                                       "Re-instantiating it as a ""used"" Common Component will requires the following steps:" & vbLf & _
                                       "Remove the component > Save the Workbook > Re-Import it from the ""Common-Components"" folder."
                    BttnUpdate = "Update"
                    sUpdateBttnTxt = "The outdated ""used"" Common Component becomes up-to-date again."
                    Set cllButtons = mMsg.Buttons(BttnUpdate, mDiff.ServicedExportVersusPublicBttn, vbLf, BttnSkipForNow, BttnSkipForever)
                Case Else
                    Stop
            End Select
            
        End With
        
        mCompMan.MsgInit
        i = 0
        With Msg
            .Section(NextSect(i)).Text.Text = sUpdateNote
            With .Section(NextSect(i))
                .Label.Text = "Last modified:"
                .Text.Text = "In Workbook : " & Comp.PublicLastModIn & vbLf & _
                             "By User     : " & Comp.PublicLastModBy & vbLf & _
                             "On Computer : " & Comp.PublicLastModOn & vbLf & _
                             "At Date/Time: " & Comp.PublicLastModAt
                .Text.MonoSpaced = True
            End With
            
            With .Section(NextSect(i))
                .Label.Text = "Update:"
                .Text.Text = sUpdateBttnTxt
            End With
            With .Section(NextSect(i))
                .Label.Text = "Display difference:"
                .Text.Text = "The displayed code modifications (return with Esc) may help estimating the extent to which re-testing is appropriate after an update."
            End With
            With .Section(NextSect(i))
                .Label.Text = "Skip for now:"
                .Text.Text = sSkipForNow
            End With
            With .Section(NextSect(i)).Text
                .Text = sSkipForNowNote
                .FontBold = True
            End With
            If CommonServiced.KindOfComponent(sComp) = enCompCommonUsed Then
                '~~ Skip forever is an option only for used Common Components (hosted need to be de-registered as such first)
                With .Section(NextSect(i))
                    .Label.Text = "Skip forever:"
                    .Text.Text = sSkipForeverNote
                End With
            End If
            With .Section(NextSect(i))
                With .Label
                    .Text = "About:"
                    .OnClickAction = GITHUB_REPO_URL & "#common-components"
                    .FontUnderline = True
                End With
                .Text.Text = mCompMan.AboutCommComps
            End With
        End With
        
        Do
            Select Case mMsg.Dsply(dsply_title:=UpdateTitle _
                                 , dsply_msg:=Msg _
                                 , dsply_Label_spec:="R70" _
                                 , dsply_buttons:=cllButtons _
                                 , dsply_modeless:=False _
                                 , dsply_width_min:=70 _
                                 , dsply_height_max:=85 _
                                 , dsply_pos:=UpdateDialogTop & ";" & UpdateDialogLeft)
                
                Case BttnUpdate:                            ChoiceUpdate sComp
                                                            Prgrss.ItemDone = sComp
                                                            Exit Do
                Case mDiff.ServicedExportVersusPublicBttn:  mDiff.ServicedExportVersusPublicDsply Comp
                Case BttnSkipForNow:                        ChoiceSkipForNow sComp
                                                            Prgrss.ItemSkipped
                                                            Exit Do
                Case BttnSkipForever:                       ChoiceSkipForever sComp
                                                            Prgrss.ItemSkipped
                                                            Exit Do
            End Select
        Loop
        Qoutdated.DeQueue
    Loop
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ChoiceSkipForever(ByVal u_comp As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "ChoiceSkipForever"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    CommonServiced.KindOfComponent(u_comp) = enCompCommonPrivate
    
    With Services
        .NoOfItemsSkipped = .NoOfItemsSkipped + 1
        .Log(u_comp) = "Outdated used Commpon Component: Update skipped forever!"
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ChoiceSkipForNow(ByVal u_comp As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "ChoiceSkipForNow"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    With Services
        .NoOfItemsSkipped = .NoOfItemsSkipped + 1
        .Log(u_comp) = "Outdated used Commpon Component: Update skipped for now!"
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub UpdateCompManDat(ByVal u_comp As String)
' ------------------------------------------------------------------------------
' Update the properties in the CommComps.dat Private Profile file with the
' properties from the CommComp.dat Private Profile file.
' ------------------------------------------------------------------------------

    With CommonServiced
        .LastModAt(u_comp) = CommonPublic.LastModAt(u_comp)
        .LastModBy(u_comp) = CommonPublic.LastModBy(u_comp)
        .LastModInWrkbkFullName(u_comp) = CommonPublic.LastModInWrkbkFullName(u_comp)
        .LastModOn(u_comp) = CommonPublic.LastModOn(u_comp)
    End With

End Sub

Private Sub ChoiceUpdate(ByVal o_comp As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "ChoiceUpdate"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    
    mUpdate.ByReImport b_comp_name:=o_comp _
                     , b_export_file:=CommonPublic.LastModExpFile(o_comp) _
                     , b_monitor:=False
    UpdateCompManDat o_comp
    
    '~~ Update the properties in the CommComps.dat file with those from the CommComps.dat file
    With New clsComp
        .CompName = o_comp
        .SetServicedEqualPublic
    End With
    
    With Services
        .NoOfItemsServicedIncrement ' = .NoOfItemsServiced + 1
        .NoOfItemsServicedNames = o_comp
        .Progress "Common Components updated"
        .Log(o_comp) = "Serviced Common Component updated by re-import of the public Export-File from the Common-Components folder."
    End With
    
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mCommComps" & "." & es_proc
End Function

Public Sub Update()
' ------------------------------------------------------------------------------
' Collects all used outdated Common Components when called for the first time
' and displays the first one queued in Qoutdated. The service is re-called until
' the queue is empty. The display of the update choices is a mode-less dialog
' which calls sub-services in accordance with the button pressed.
' ------------------------------------------------------------------------------
    Const PROC = "Update"
    
    On Error GoTo eh
    
    CollectOutdated
    ChoiceLoop
    CommonPending.Info
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub CollectOutdated(Optional ByRef c_outdated As clsQ)
' ------------------------------------------------------------------------------
' Collects all outdated Used Common Components and enqueues them in Qoutdated.
' For any up-to-date Common Component the last update datetime is maintained in
' the CommComps.dat Private Profile file.
' ------------------------------------------------------------------------------
    Const PROC = "CollectOutdated"
    
    On Error GoTo eh
    Dim Comp            As clsComp
    Dim wbk             As Workbook
    Dim dct             As Dictionary
    Dim v               As Variant
    Dim sComp           As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.ServicedWbk
    Set Qoutdated = New clsQ
    Set dct = Serviced.CompsCommon
    Prgrss.ItemsTotal = dct.Count
    Prgrss.Operation = "outdated"
    
    For Each v In dct
        sComp = v
        Set Comp = New clsComp
        With Comp
            .CompName = sComp
            Services.LogItem = Serviced.Wrkbk.VBProject.VBComponents(.CompName)
            Select Case CommonServiced.KindOfComponent(sComp)
                Case enCompCommonHosted, enCompCommonUsed
                    If .CodeCrrent.Meets(.CodePublic) = False _
                    And Not .CommCompIsPendingByServicedWorkbook Then ' outdated
                        Qoutdated.EnQueue Comp
                        Prgrss.ItemDone = sComp
                        With Services
                            .NoOfItemsServicedNames = sComp
                            .NoOfItemsOutdated = Qoutdated.Size
                        End With
                    Else
                        UpdateCompManDat sComp
                        Prgrss.ItemSkipped
                        If .ServicedLastModAt < .PublicLastModAt Then
                            .ServicedLastModAt = .PublicLastModAt
                        End If
                        With Services
                            .NoOfItemsSkipped = .NoOfItemsSkipped + 1
                            .Log(sComp) = "Serviced Common Component used is up-to-date"
                        End With
                    End If
                Case Else
                    Prgrss.ItemSkipped
            End Select
        End With
        Set Comp = Nothing
        Application.StatusBar = vbNullString
        DoEvents
    Next v
    Prgrss.Dsply
    
xt: Set c_outdated = Qoutdated
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

