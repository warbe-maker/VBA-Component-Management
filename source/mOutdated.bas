Attribute VB_Name = "mOutdated"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mOutdated Procedures for the 'UpdateOutdatedCommonComponents
'                           service.
' Public services:
'
' ----------------------------------------------------------------------------

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mOutdated." & sProc
End Function

Public Sub Display(ByVal d_hosted As String)
' ----------------------------------------------------------------------------
' Displays all (still) uotdated Common Components in the serviced Workbook.
' ----------------------------------------------------------------------------
    Const PROC = "Display"
    
    On Error GoTo eh
    Dim cllBttns    As New Collection
    Dim Comp        As clsComp
    Dim dctOutdated As Dictionary
    Dim dctRunArgs  As Dictionary
    Dim fUpdate     As fMsg
    Dim i           As Long
    Dim Msg         As TypeMsg
    Dim sBttn1      As String
    Dim sBttn2      As String
    Dim sTitle      As String
    
    mBasic.BoP ErrSrc(PROC)
    sTitle = "To-be-updated outdated Common Component(s)"
    mMsg.MsgInstance sTitle, True ' unload any previous
    
    If mOutdated.Collect(dctOutdated).Count = 0 Then
        mService.DsplyStatus Log.Service & " Done! (" & wsService.CommonComponentsUpdated & " of " & _
                             wsService.CommonComponentsUsed & " used Common Components updated/renewed)"
        GoTo xt
    End If
    
    mMsg.MsgInstance sTitle
    '~~ Prepare a modeless message with a pair of buttons for each outdated component.
    '~~ The mMsg.Dsply service allows only 7 button rows. The number of rows might
    '~~ thus not cover all outdated components. Any exessive components will be
    '~~ displayed by subsequent calls of this service.
    Set fUpdate = mMsg.MsgInstance(sTitle)
    For i = 1 To mBasic.Min(dctOutdated.Count, 7)
        Set Comp = dctOutdated.Items()(i - 1)
        With Comp
            sBttn1 = .CompName & vbLf & vbLf & "Update"
            sBttn2 = .CompName & vbLf & vbLf & "Changes"
            mMsg.ButtonAppRun dctRunArgs, sBttn1, ThisWorkbook, "mOutdated.RenewByReImport", .Wrkbk, .CompName, .Raw.SavedExpFileFullName, d_hosted
            mMsg.ButtonAppRun dctRunArgs, sBttn2, ThisWorkbook, "mService.ExpFilesDiffDisplay", .ExpFileFullName, .Raw.SavedExpFileFullName, "Currently used (" & .ExpFileFullName & ")", "Up-to-date (" & .Raw.SavedExpFileFullName & ")"
        End With
        Set cllBttns = mMsg.Buttons(cllBttns, sBttn1, sBttn2, vbLf)
    Next i
    With Msg.Section(1)
        .Label.Text = "Update:"
        .Label.FontColor = rgbDarkGreen
        .Text.Text = "Press/click the button to update the desired outdated Common Component. The component " & _
                     "will be updated and the dialog re-displayed - without the then already updated component, " & _
                     "until there's no outdated component left or the dialog is closed explicitely."
    End With
    With Msg.Section(2)
        .Label.Text = "Changes:"
        .Label.FontColor = rgbDarkGreen
        .Text.Text = "Press/click the button to display/check which of the code has changed"
    End With
    With Msg.Section(3)
        .Label.Text = "About:"
        .Label.FontColor = rgbDarkGreen
        .Text.Text = "Experience has shown that this way of renewing outdated Common Components. " & _
                     "I.e. updating one component at a time by a service performed via Application.Run " & _
                     "is the most stable approach. Since the dialog is displayed modeless the serviced " & _
                     "workbook may be saved after each individual update - and will be saved prior an update."
    End With
    
    '~~ Display a modeless message with a pair of buttons for each outdated component
    mMsg.Dsply dsply_title:=sTitle _
             , dsply_msg:=Msg _
             , dsply_buttons:=cllBttns _
             , dsply_modeless:=True _
             , dsply_buttons_app_run:=dctRunArgs _
             , dsply_width_min:=40
                    
xt: Set dctOutdated = Nothing
    Set Log = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function Collect(ByRef c_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the serviced Workbook's (still) outdated Common Components (c_dct) as
' Dictionary.
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim dctAll          As Dictionary
    Dim dctOutdated     As New Dictionary
    Dim vbc             As VBComponent
    Dim fso             As New FileSystemObject
    Dim sOutdated       As String
    Dim v               As Variant
    Dim Comp            As clsComp
    Dim Comps           As New clsComps
    Dim lAll            As Long
    Dim lRemaining      As Long
    Dim wbk             As Workbook
    Dim lUsed           As Long
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = mService.WbkServiced
    Set dctAll = mService.AllComps(wbk)

    With wbk.VBProject
        lAll = .VBComponents.Count
        lRemaining = lAll
        For Each v In dctAll
            Set vbc = dctAll(v)
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = wbk
                .CompName = vbc.name
                Set .VBComp = vbc
                If .KindOfComp = enCommCompUsed Then
                    lUsed = lUsed + 1
                    If .Outdated Then
                        mService.AddAscByKey dctOutdated, .CompName, Comp
                        sOutdated = .CompName
                    End If ' .Outdated
                End If ' Used Common Component
            End With
            Set Comp = Nothing
            lRemaining = lRemaining - 1
            mService.DsplyStatus _
            mService.Progress(p_service:=Log.Service _
                            , p_result:=dctOutdated.Count _
                            , p_of:=lAll _
                            , p_op:="used Common Components are outdated" _
                            , p_comps:=sOutdated _
                            , p_dots:=lRemaining _
                             )
        Next v
    End With
    
    Set Collect = dctOutdated
    Set c_dct = dctOutdated
    mService.DsplyStatus _
    mService.Progress(p_service:=Log.Service _
                    , p_result:=dctOutdated.Count _
                    , p_of:=lAll _
                    , p_op:="used Common Components are outdated" _
                    , p_comps:=sOutdated _
                     )
    
    
xt: If wsService.CommonComponentsUsed = 0 Then wsService.CommonComponentsUsed = lUsed
    If wsService.CommonComponentsOutdated = 0 Then wsService.CommonComponentsOutdated = dctOutdated.Count
    Set dctOutdated = Nothing
    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub RenewByReImport(ByRef rbi_wbk_serviced As Workbook, _
                           ByVal rbi_vbc_name As String, _
                           ByVal rbi_exp_file As String, _
                  Optional ByVal rbi_hosted As String = vbNullString)
' ------------------------------------------------------------------------------
' Called via Application.Run by a CommonButton of the message displaying all
' still outdated Common Coponents: Renews a Workbook's (rbi_wbk_serviced) 'Common
' Component' (rbi_vbc_name) by re-importing the up-to-date component's export file
' (rbi_exp_file) from the Common Components folder. After execution the Display
' service is called again to redisplay still outstanding to be updated outdated
' 'Common Components'.
'
' Preconditions (not evaluated by the service):
' - rbi_wbk_serviced is an open Workbook
' - rbi_vbc_name is a components in its VB-Prtoject
' - rbi_exp_file_full-Name is a file in the "Common Components" folder
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------
    Const PROC              As String = "RenewByReImport"
    Const MONITORED_STEPS   As Long = 11
    Const MONITOR_POS       As String = "20;400"
    
    On Error GoTo eh
    Dim sTempName       As String
    Dim fso             As New FileSystemObject
    Dim Comp            As clsComp
    Dim TmpWbk          As Workbook
    Dim KoC             As Long
    Dim MonitorFooter   As TypeMsgText
    Dim MonitorStep     As TypeMsgText
    Dim MonitorTitle    As String
    Dim Step            As Long
    
    mBasic.BoP ErrSrc(PROC)
    MonitorTitle = "Update an outdated component"
    MonitorFooter.FontColor = rgbDarkGreen
    MonitorFooter.FontSize = 9
    mMsg.MsgInstance MonitorTitle, True ' close any previous monitor window
    mMsg.MsgInstance MonitorTitle       ' establish a new monitor window
    MonitorStep.MonoSpaced = True
    MonitorStep.FontSize = 9
    Set Log = New clsLog
    Log.Service(True) = mCompManClient.SRVC_UPDATE_OUTDATED ' new log file when older 1 day
    
    '~~ Save the serviced Workbook when yet not saved (initialize the monitor window)
    If Not rbi_wbk_serviced.Saved Then
        Step = Step + 1
        MonitorStep.Text = Step & ". Save serviced Workbook '" & rbi_wbk_serviced.name & "'"
        mMsg.Monitor mon_title:=MonitorTitle _
                   , mon_text:=MonitorStep _
                   , mon_steps_displayed:=MONITORED_STEPS _
                   , mon_width_min:=70 _
                   , mon_pos:=MONITOR_POS
'        mService.WbkSave rbi_wbk_serviced
    End If
    
    Set Comp = New clsComp
    With Comp
        Set .Wrkbk = rbi_wbk_serviced
        .CompName = rbi_vbc_name
        Log.ServicedItem = .VBComp
    End With
    EstablishExecTraceFile rbi_wbk_serviced
    
    '~~ Create and activate a hidden Workbook
    '~~ Because this may not be the first step the monitor initialization values are provided.
    '~~ They are ignored by the service when the monitor window is already initialized
    Step = Step + 1
    MonitorStep.Text = Step & ". Activate a hidden temporary Workbook."
    mMsg.Monitor mon_title:=MonitorTitle _
               , mon_text:=MonitorStep _
               , mon_steps_displayed:=MONITORED_STEPS _
               , mon_width_min:=70 _
               , mon_pos:=MONITOR_POS
    Set TmpWbk = TempWbkHidden()
    TmpWbk.Activate
              
    With rbi_wbk_serviced.VBProject
        If mComp.Exists(xst_wbk:=rbi_wbk_serviced, xst_vbc_name:=rbi_vbc_name) Then
        
            '~~ Rename an already existing component
            sTempName = mComp.TempName(tn_wbk:=rbi_wbk_serviced, tn_vbc_name:=rbi_vbc_name)
            Step = Step + 1
            MonitorStep.Text = Step & ". Rename the component '" & rbi_vbc_name & "' to '" & sTempName & "'."
            mMsg.Monitor MonitorTitle, MonitorStep
            .VBComponents(rbi_vbc_name).name = sTempName
            
            '~~ Outcomment the renamed component's code
            Step = Step + 1
            MonitorStep.Text = Step & ". Outcomment all code lines in the renamed component."
            mMsg.Monitor MonitorTitle, MonitorStep
            OutCommentCodeInRenamedComponent rbi_wbk_serviced, sTempName ' this had made it much less "reliablele"
            mBasic.TimedDoEvents ErrSrc(PROC)
            
            '~~ Remove the renamed component (postponed thought)
            Step = Step + 1
            MonitorStep.Text = Step & ". Remove the renamed component."
            mMsg.Monitor MonitorTitle, MonitorStep
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            
        End If
        
        '~~ (Re-)import the component
        Step = Step + 1
        MonitorStep.Text = Step & ". (Re-) import the Export File of the up-to-date version of the component."
        mMsg.Monitor MonitorTitle, MonitorStep
        .VBComponents.Import rbi_exp_file
        Log.Entry = "'" & rbi_vbc_name & "' (re-)imported from '" & rbi_exp_file & "'"
                
        '~~ Export the re-newed Used Common Component
        Step = Step + 1
        MonitorStep.Text = Step & ". Export the (re-)imported component."
        mMsg.Monitor MonitorTitle, MonitorStep
        Set Comp = New clsComp
        With Comp
            Set Comp.Wrkbk = rbi_wbk_serviced
            .CompName = rbi_vbc_name
        End With
        .VBComponents(rbi_vbc_name).Export Comp.ExpFileFullName
        
        '~~ Remove the activated hidden Workbook and re-activate the serviced Workbook
        Step = Step + 1
        MonitorStep.Text = Step & ". Remove the temporary activated and re-activate the serviced Workbook."
        mMsg.Monitor MonitorTitle, MonitorStep
        TempWbkHiddenRemove TmpWbk
        rbi_wbk_serviced.Activate
  
        '~~ Set the updated component's RevisionNumber = the RevisionNumber of the re-imported component's Export file
        Step = Step + 1
        MonitorStep.Text = Step & ". Update the updated component's revision number."
        mMsg.Monitor MonitorTitle, MonitorStep
        With Comp
            Set Comp.Wrkbk = rbi_wbk_serviced
            .CompName = rbi_vbc_name
            KoC = .KindOfComp ' this establishes the .Raw instance
            Log.Entry = "Used Common Component renewed/updated by (re-)import of the Raw's Export File (" & .Raw.SavedExpFileFullName & ")"
            .RevisionNumber = .Raw.RevisionNumber
            .DueModificationWarning = False
            .Export
        End With

        '~~ Re-call the update service in order to display still outdated components if any
        MonitorFooter.Text = "Outdated component '" & rbi_vbc_name & "' successfully updated."
        mMsg.MonitorFooter MonitorTitle, MonitorFooter
        mOutdated.Display rbi_hosted ' display still outdated commen components mode-less
    
        wsService.CommonComponentsUpdated = wsService.CommonComponentsUpdated + 1
    End With
    
xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function TempWbkHidden() As Workbook
    Dim app As Excel.Application
    
    Set app = CreateObject("Excel.Application")
    app.Visible = False
    Set TempWbkHidden = app.Workbooks.Add()

End Function

Public Sub OutCommentCodeInRenamedComponent(ByRef oc_wbk As Workbook, _
                                            ByVal oc_component As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutCommentCodeInRenamedComponent"
    
    On Error GoTo eh
    Dim i As Long
    
    mBasic.BoP ErrSrc(PROC)
    With oc_wbk.VBProject.VBComponents(oc_component).CodeModule
        For i = 1 To .CountOfLines
            .ReplaceLine i, "'" & .Lines(i, 1)
        Next i
    End With

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub TempWbkHiddenRemove(ByVal wbk As Workbook)
    Dim app As Excel.Application
    
    Set app = wbk.Parent
    wbk.Close
    Set app = Nothing

End Sub

