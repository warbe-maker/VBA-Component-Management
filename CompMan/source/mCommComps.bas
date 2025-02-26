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
Public Const BTTN_UPDATE    As String = "Update"
Public Const BTTN_Terminate As String = "Terminate"
Private BttnSkipForever     As String
Private BttnSkipForNow      As String
Private Qoutdated           As clsQ
Private UpdateDialogLeft    As Long
Private UpdateDialogTop     As Long
Private UpdateTitle         As String

Private Sub ChoiceLoop()
' ------------------------------------------------------------------------------
' Loops through enqued outdate Common Component displaying a dialog with the
' choices: update, display diffs, skip for now, and skip forever.
' ------------------------------------------------------------------------------
    Const PROC = "ChoiceLoop"
    
    On Error GoTo eh
    Dim cllButtons          As Collection
    Dim Comp                As clsComp
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
                    sUpdateBttnTxt = "With this update the ""hosted"" Common Component will become in sync with all Workbook using this version."
                    Set cllButtons = mMsg.Buttons(BTTN_UPDATE, mDiff.ServicedExportVersusPublicBttn, vbLf, BttnSkipForNow)
    
                Case enCompCommonUsed
                    UpdateTitle = "Used ""Common Component"" is not/no longer up-to-date!"
                    sUpdateNote = "The used Common Component   " & mBasic.Spaced(sComp) & "   is outdated."
                    sSkipForNow = "Update will be postponed and proposed with the next Workbook open. While not up-to-date " & _
                                  "any code modification will be discarded with any subsequent export service."
                    sSkipForeverNote = "The component, currently known as a used ""Common Component"", will be de-registered " & _
                                       "and no longer updated in the future (its ""used"" status will be changed into ""private""). " & _
                                       "Re-instantiating it as a ""used"" Common Component will requires the following steps:" & vbLf & _
                                       "Remove the component > Save the Workbook > Re-Import it from the ""Common-Components"" folder."
                    sUpdateBttnTxt = "The outdated ""used"" Common Component becomes up-to-date again."
                    Set cllButtons = mMsg.Buttons(BTTN_UPDATE, mDiff.ServicedExportVersusPublicBttn, vbLf, BttnSkipForNow, BttnSkipForever, BTTN_Terminate)
                Case Else
                    Stop
            End Select
            
        End With
        
        mCompMan.MsgInit
        i = 0
        NextSect i
        Msg.Section(i).Text.Text = sUpdateNote
        NextSect i
        Msg.Section(i).Label.Text = "Last modified:"
        With Comp
            Msg.Section(i).Text.Text = "In Workbook : " & .PublicLastModIn & vbLf & _
                                       "By User     : " & .PublicLastModBy & vbLf & _
                                       "On Computer : " & .PublicLastModOn & vbLf & _
                                       "At Date/Time: " & .PublicLastModAt
        End With
        Msg.Section(i).Text.MonoSpaced = True
            
        With Msg
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
            Select Case mMsg.Dsply(d_title:=UpdateTitle _
                                 , d_msg:=Msg _
                                 , d_label_spec:="R70" _
                                 , d_buttons:=cllButtons _
                                 , d_modeless:=False _
                                 , d_width_min:=70 _
                                 , d_height_max:=85 _
                                 , d_pos:=UpdateDialogTop & ";" & UpdateDialogLeft)
                
                Case BTTN_UPDATE:                           ChoiceUpdate sComp
                                                            Prgrss.ItemDone = sComp
                                                            Exit Do
                Case mDiff.ServicedExportVersusPublicBttn:  mDiff.ServicedExportVersusPublicDsply Comp
                Case BttnSkipForNow:                        ChoiceSkipForNow sComp
                                                            Prgrss.ItemSkipped
                                                            Exit Do
                Case BttnSkipForever:                       ChoiceSkipForever sComp
                                                            Prgrss.ItemSkipped
                                                            Exit Do
                Case BTTN_Terminate:                        Qoutdated.Clear
                                                            GoTo xt
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
    
    With Servicing
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
    With Servicing
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

Private Sub ChoiceUpdate(ByVal o_comp As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "ChoiceUpdate"
    
    On Error GoTo eh
    Dim bDone As Boolean
    
    mBasic.BoP ErrSrc(PROC)
    mUpdate.ByReImport b_comp_name:=o_comp _
                     , b_export_file:=CommonPublic.LastModExpFile(o_comp) _
                     , b_monitor:=False
    
    '~~ Update the properties in the CommComps.dat file with those from the CommComps.dat file
    CommonServiced.SetPropertiesEqualPublic o_comp
    
    With Servicing
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

Public Sub CollectOutdated(Optional ByRef c_outdated As clsQ)
' ------------------------------------------------------------------------------
' Collects all outdated Used Common Components and enqueues them in Qoutdated.
' For any up-to-date Common Component the last update datetime is maintained in
' the CommComps.dat Private Profile file.
' ------------------------------------------------------------------------------
    Const PROC = "CollectOutdated"
    
    On Error GoTo eh
    Dim bAny    As Boolean
    Dim bDone   As Boolean
    Dim Comp    As clsComp
    Dim wbk     As Workbook
    Dim dct     As Dictionary
    Dim v       As Variant
    Dim sComp   As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Servicing.ServicedWbk
    Set Qoutdated = New clsQ
    Set dct = Serviced.CompsCommon
    Prgrss.ItemsTotal = dct.Count
    Prgrss.Operation = "outdated"
    
    With Serviced
        For Each v In dct
            sComp = v
            .Comp.CompName = sComp
            Servicing.LogItem = .Wrkbk.VBProject.VBComponents(sComp)
            Select Case True
                Case CommonServiced.KindOfComponent(sComp) = enCompInternal, CommonServiced.KindOfComponent(sComp) = enCompCommonPrivate
                    '~~ Ignored
                Case mCommComps.IsPendingByServicedWorkbook(.Comp)
                    '~~ Ignored while pending release
                    Prgrss.ItemSkipped
                Case mDiff.ServicedCodeVersusPublic(.Comp) = False
                    '~~ Used or hosted, not pending release, not outdated
                    '~~ Just make sure the origin is correctly documented
                    CommonServiced.SetPropertiesEqualPublic sComp
                    With Servicing
                        .NoOfItemsSkipped = .NoOfItemsSkipped + 1
                        .Log(sComp) = "Serviced Common Component used is up-to-date"
                    End With
                    Prgrss.ItemSkipped
                Case Else
                    Set Comp = New clsComp
                    Comp.CompName = sComp
                    Qoutdated.EnQueue Comp
                    Prgrss.ItemDone = sComp
                    With Servicing
                        .NoOfItemsServicedNames = sComp
                        .NoOfItemsOutdated = Qoutdated.Size
                        Debug.Print "Queue " & Qoutdated.Size & ": " & sComp
                    End With
            End Select
            Application.StatusBar = vbNullString
            DoEvents
        Next v
    End With
    Prgrss.Dsply
    
xt: Set c_outdated = Qoutdated
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mCommComps" & "." & es_proc
End Function

Public Function HasModificationPendingRelease(ByVal p_comp As String, _
                                     Optional ByRef p_last_mod_at_datetime_utc As String, _
                                     Optional ByRef p_last_mod_export_filename As String, _
                                     Optional ByRef p_last_mod_in_wbk_fullname As String, _
                                     Optional ByRef p_last_mod_on_machine As String) As Boolean
    
    If CommonPending Is Nothing Then
        Set CommonPending = New clsCommonPending
        CommonPending.CompName = p_comp
    End If
    HasModificationPendingRelease = CommonPending.Exists(p_comp _
                                                       , p_last_mod_at_datetime_utc _
                                                       , p_last_mod_export_filename _
                                                       , p_last_mod_in_wbk_fullname _
                                                       , p_last_mod_on_machine)
    
End Function

Public Function IsPendingByServicedWorkbook(ByVal i_comp As clsComp)

    Dim sLastModInWbkFullName As String

    With i_comp
        If mCommComps.HasModificationPendingRelease(.CompName, , , sLastModInWbkFullName) Then
            IsPendingByServicedWorkbook = sLastModInWbkFullName = Serviced.Wrkbk.FullName _
                                              And .CodeExprtd.Meets(.CodePnding)
        End If
    End With

End Function

Private Function NextSect(ByRef n_sect As Long) As Long
    n_sect = n_sect + 1
    NextSect = n_sect
End Function

Public Sub StatusOfCommComp(ByVal s_comp As String)
    Const PROC = "StatusOfCommComp"
    
    On Error GoTo eh
    mCompMan.ServiceInitiate s_serviced_wbk:=ActiveWorkbook _
                           , s_service:="Release Common Components" _
                           , s_do_housekeeping:=False
    
    If CommonServiced Is Nothing Then Set CommonServiced = New clsCommonServiced
    With CommonServiced
        If .PPFile.Exists(s_comp) Then
            Debug.Print "Serviced: Last mod at = .Properties.LmAt"
        Else
            Debug.Print "Serviced: n o t  e x i s t s !"
        End If
    End With
    
    If CommonPending Is Nothing Then Set CommonPending = New clsCommonPending
    With CommonPending
        If .Exists(s_comp) Then
            Debug.Print "Pending : Last mod at = .Properties.LmAt"
        Else
            Debug.Print "Pending : n o t  e x i s t s !"
        End If
    End With
    
    If CommonPublic Is Nothing Then Set CommonPublic = New clsCommonPublic
    With CommonPublic
        If .Exists(s_comp) Then
            Debug.Print "Public  : Last mod at = .Properties.LmAt"
        Else
            Debug.Print "Public  : n o t  e x i s t s !"
        End If
    End With
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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

