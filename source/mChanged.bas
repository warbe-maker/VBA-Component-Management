Attribute VB_Name = "mChanged"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mChanged
' Procedures and services for the 'Update Outdated Common Components' service
' and the 'Synchronize VB-Project' service.
'
' Public services:
' - UpdateByReplacingAllCodeLines
'   Synchronizes the code of a Sync-Target-Workbook's VBComponent by
'   replacing the code lines with the code lines of the corrresponding Sync-
'   Source-Workbook's VBComponent.
'
' - CollectCompsTheCodeDiffers
'   Returns a collection of those VBComponents in the Sync-Target-Workbook
'   whith a different code in the corresponding VBComponent in the Sync-
'   Source-Workbook.

'Public Function CollectCompsTheCodeDiffers() As Dictionary
'' ------------------------------------------------------------------------------
'' Returns a collection of those VBComponents in the Sync-Target-Workbook
'' (sync_target_workbook) whith a different code in the corresponding VBComponent
'' in the Sync-Source-Workbook.
'' ------------------------------------------------------------------------------
'    Const PROC = "CollectCompsTheCodeDiffers"
'
'    On Error GoTo eh
'    Dim cll         As Collection
'    Dim dct         As New Dictionary
'    Dim SourceComp  As clsComp
'    Dim TargetComp  As clsComp
'    Dim v           As Variant
'    Dim vbc         As VBComponent
'    Dim sProgress   As String
'
'    mBasic.BoP ErrSrc(PROC)
'    mSync.MonitorStep "Collecting VB-Components the code has changed"
'
'    sProgress = String(mSync.Source.VBProject.VBComponents.Count, ".")
'    For Each vbc In mSync.Source.VBProject.VBComponents
'        mSync.MonitorStep "Collecting VB-Components the code has changed " & sProgress
'        Set SourceComp = New clsComp
'        With SourceComp
'            Set .Wrkbk = mSync.Source
'            .CompName = vbc.Name
'        End With
'        If mComp.Exists(vbc, mSync.TargetWorkingCopy) Then
'            Set TargetComp = New clsComp
'            With TargetComp
'                Set .Wrkbk = mSync.TargetWorkingCopy
'                .CompName = vbc.Name
'            End With
'            If CompChanged(SourceComp.CodeLines, TargetComp.CodeLines) Then
'                mService.Log.ServicedItem = vbc
'                Set cll = New Collection
'                cll.Add mComp.TypeString(vbc) & _
'                        " " & vbc.Name & _
'                        vbLf & vbLf & _
'                        "Synchronize code change"   ' 1. The button's caption
'                cll.Add ThisWorkbook                ' 2. The servicing Workbook
'                cll.Add "mSyncComps.AppRunChanged"    ' 3. The service to run
'                cll.Add vbc.Name                    ' 4. The VBComponent to update/renew
'                mDct.DctAdd dct, mSyncComps.SyncId(vbc), cll, order_bykey, seq_ascending, sense_casesensitive
'                Set cll = Nothing
'            End If
'            Set TargetComp = Nothing
'            Set SourceComp = Nothing
'        End If
'        sProgress = Left(sProgress, Len(sProgress) - 1)
'    Next vbc
'
'    mSync.MonitorStep "Collecting VB-Components the code has changed " & sProgress
'
'xt: Set CollectCompsTheCodeDiffers = dct
'    Set dct = Nothing
'    mBasic.EoP ErrSrc(PROC)
'    Exit Function
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function

Public Sub DisplayOutdated(ByVal d_hosted As String)
' ----------------------------------------------------------------------------
' Displays all (still) outdated Common Components in the serviced Workbook.
' ----------------------------------------------------------------------------
    Const PROC = "DisplayOutdated"
    
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
    
    If mChanged.CollectOutdatedCommonComps(dctOutdated).Count = 0 Then
        mService.DsplyStatus " Done! (" & wsService.CommonComponentsUpdated & " of " & _
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
            mMsg.ButtonAppRun dctRunArgs, sBttn1, ThisWorkbook, "mChanged.ReImportCommonComponent", .Wrkbk, .CompName, .Raw.SavedExpFileFullName, d_hosted
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
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mChanged." & sProc
End Function

Private Sub OutCommentCodeInRenamedComponent(ByRef oc_wbk As Workbook, _
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

Public Sub ReImport(ByRef bi_wbk_serviced As Workbook, _
                            ByVal bi_vbc_name As String, _
                            ByVal bi_exp_file As String)
' ------------------------------------------------------------------------------
' This service renews a Workbook's (bi_wbk_serviced) component (bi_vbc_name) by
' re-importing a more up-to-date component's export file (bi_exp_file). When
' the service is performed directly by Application.Run (bi_via_app_run = True)
' it calls mCompMan.Update again to redisplay the still to be updated outdated
' components.
' Preconditions (not evaluated by the service):
' - bi_wbk_serviced is an open Workbook
' - bi_vbc_name is a components in its VB-Project
' - bi_exp_file_full-Name is a file in the "Common Components" folder
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------
    Const PROC = "ReImport"
    
    On Error GoTo eh
    Dim sTempName   As String
    Dim fso         As New FileSystemObject
    Dim Comp        As clsComp
    Dim TmpWbk      As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    If Log Is Nothing Then
        '~~ The service has been called directly for a single individual update
        mService.Log.Service = mCompManClient.SRVC_UPDATE_OUTDATED
        Set Comp = New clsComp
        With Comp
            Set .Wrkbk = bi_wbk_serviced
            .CompName = bi_vbc_name
            mService.Log.ServicedItem = .VBComp
        End With
        mService.EstablishExecTraceFile bi_wbk_serviced
    End If
    
    With bi_wbk_serviced.VBProject
        
        mTrc.BoC "Create and activate a hidden Workbook"
        Set TmpWbk = TempWbkHidden()
        TmpWbk.Activate
        mTrc.EoC "Create and activate a hidden Workbook"
        
        If mComp.Exists(ex_wbk:=bi_wbk_serviced, ex_vbc:=bi_vbc_name) Then
            '~~ Rename the already existing component ...
            mTrc.BoC "Rename the to-be-removed component"
            sTempName = mComp.TempName(tn_wbk:=bi_wbk_serviced, tn_vbc_name:=bi_vbc_name)
            .VBComponents(bi_vbc_name).Name = sTempName
            mTrc.EoC "Rename the to-be-removed component"
            
            '~~ ... and outcomment its code.
            '~~ An (irregular) Workbook close may leave renamed components un-removed.
            '~~ When the Workbook is re-opened again any renamed component is deleted.
            '~~ However VB may fail when dedecting duplicate declarations before the
            '~~ renamed component can be deleted. Outcommenting the renamded component
            '~~ should prevent this.
            mChanged.OutCommentCodeInRenamedComponent bi_wbk_serviced, sTempName ' this had made it much less "reliablele"
            mBasic.TimedDoEvents ErrSrc(PROC)
        
            '~~ Remove the renamed component (postponed thought)
            mTrc.BoC "Remove the renamde component"
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            mTrc.EoC "Remove the renamde component"
        End If
        
        '~~ (Re-)import the component
        mTrc.BoC "(Re-)Import the up-to-date Raw Common Component"
        .VBComponents.Import bi_exp_file
        mService.Log.Entry = "Renewed/updated by (re-)import of '" & bi_exp_file & "'"
        mTrc.EoC "(Re-)Import the up-to-date Raw Common Component"
                
        '~~ Export the re-newed Used Common Component
        mTrc.BoC "Export the renewed Common Component"
        Set Comp = New clsComp
        With Comp
            Set Comp.Wrkbk = bi_wbk_serviced
            .CompName = bi_vbc_name
            .Export
        End With
        mTrc.EoC "Export the renewed Common Component"
        
        mTrc.BoC "Remove the activated hidden Workbook and re-activate the processed Workbook"
        TempWbkHiddenRemove TmpWbk
        bi_wbk_serviced.Activate
        mTrc.EoC "Remove the activated hidden Workbook and re-activate the processed Workbook"

    End With
        
xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ReImportCommonComponent(ByRef r_wbk As Workbook, _
                                   ByVal r_vbc_name As String, _
                                   ByVal r_exp_file As String, _
                          Optional ByVal r_hosted As String = vbNullString)
' ------------------------------------------------------------------------------
' Called via Application.Run by a CommonButton of the message displaying all
' still outdated Common Coponents. Updates/renews the Workbook's (r_wbk)
' VBComponent (r_vbc_name) by re-importing the export file (r_exp_file).
' from the Common Components folder.
' When a 'Common Component is updated/renewed (r_common_component = True) the
' service re-calls the display of the still outstanding to be updated outdated
' 'Common Components'.
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------
    Const PROC              As String = "ReImportCommonComponent"
    
    On Error GoTo eh
    mUpdate.ByReImport b_wbk_target:=r_wbk _
                     , b_vbc_name:=r_vbc_name _
                     , b_exp_file:=r_exp_file
    
xt: mBasic.EoP ErrSrc(PROC)
    mChanged.DisplayOutdated r_hosted ' display still outdated commen components mode-less
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub UpdateByReImportSourceExportFile(ByVal u_wbk_target As Workbook, _
                                            ByVal u_vbc_name As String, _
                                            ByVal u_exp_file As String, _
                                   Optional ByVal u_monitor As Boolean = False)
' ------------------------------------------------------------------------------
' Updates a target component of which the source has changed by re-importing the
' source component's export-file to the target VB-Project.
' Preconditions:
' - The source Workbook is open
' - The VBComponent (u_vbc_name) is a components in the source Workbook's VB-Project
' - The Export-file (u_exp_file is from the source VBComponent
'
' W. Rauschenberger Berlin, Jan 2023
' ------------------------------------------------------------------------------
    Const PROC              As String = "UpdateByReImportSourceExportFile"
    Const MONITORED_STEPS   As Long = 11
    Const MONITOR_POS       As String = "20;400"
    
    On Error GoTo eh
    Dim sTempName       As String
    Dim Comp            As clsComp
    Dim TmpWbk          As Workbook
    Dim MonitorFooter   As TypeMsgText
    Dim MonitorStep     As TypeMsgText
    Dim MonitorTitle    As String
    Dim Step            As Long
    
    mBasic.BoP ErrSrc(PROC)
    If u_monitor Then
        MonitorTitle = "Update an outdated component"
        MonitorFooter.FontColor = rgbDarkGreen
        MonitorFooter.FontSize = 9
        mMsg.MsgInstance MonitorTitle, True ' close any previous monitor window
        mMsg.MsgInstance MonitorTitle       ' establish a new monitor window
        MonitorStep.MonoSpaced = True
        MonitorStep.FontSize = 9
    End If
    
    Set Comp = New clsComp
    With Comp
        Set .Wrkbk = u_wbk_target
        .CompName = u_vbc_name
    End With
    
    '~~ Create and activate a hidden Workbook
    '~~ Because this may not be the first step the monitor initialization values are provided.
    '~~ They are ignored by the service when the monitor window is already initialized
    If u_monitor Then
        Step = Step + 1
        MonitorStep.Text = Step & ". Activate a hidden temporary Workbook."
        mMsg.Monitor mon_title:=MonitorTitle _
                   , mon_text:=MonitorStep _
                   , mon_steps_displayed:=MONITORED_STEPS _
                   , mon_width_min:=70 _
                   , mon_pos:=MONITOR_POS
    End If
    Set TmpWbk = TempWbkHidden()
    TmpWbk.Activate
              
    With u_wbk_target.VBProject
        '~~ Rename an already existing component
        sTempName = mComp.TempName(tn_wbk:=u_wbk_target, tn_vbc_name:=u_vbc_name)
        If u_monitor Then
            Step = Step + 1
            MonitorStep.Text = Step & ". Rename the component '" & u_vbc_name & "' to '" & sTempName & "'."
            mMsg.Monitor MonitorTitle, MonitorStep
        End If
        .VBComponents(u_vbc_name).Name = sTempName
        
        '~~ Outcomment the renamed component's code
        If u_monitor Then
            Step = Step + 1
            MonitorStep.Text = Step & ". Outcomment all code lines in the renamed component."
            mMsg.Monitor MonitorTitle, MonitorStep
        End If
        OutCommentCodeInRenamedComponent u_wbk_target, sTempName ' this had made it much less "reliablele"
        mBasic.TimedDoEvents ErrSrc(PROC)
        
        '~~ Remove the renamed component (postponed thought)
        If u_monitor Then
            Step = Step + 1
            MonitorStep.Text = Step & ". Remove the renamed component."
            mMsg.Monitor MonitorTitle, MonitorStep
        End If
        .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
                
        '~~ (Re-)import the component
        If u_monitor Then
            Step = Step + 1
            MonitorStep.Text = Step & ". (Re-) import the Export File of the up-to-date version of the component."
            mMsg.Monitor MonitorTitle, MonitorStep
        End If
        .VBComponents.Import u_exp_file
                        
        '~~ Remove the activated hidden Workbook and re-activate the serviced Workbook
        If u_monitor Then
            Step = Step + 1
            MonitorStep.Text = Step & ". Remove the temporary activated and re-activate the serviced Workbook."
            mMsg.Monitor MonitorTitle, MonitorStep
        End If
        TempWbkHiddenRemove TmpWbk
        u_wbk_target.Activate
          
        If u_monitor Then
            MonitorFooter.Text = "Successfully updated! (process monitor closes in 2 seconds)"
            mMsg.MonitorFooter MonitorTitle, MonitorFooter
            Application.Wait Now() + TimeValue("0:00:02")
            mMsg.MsgInstance MonitorTitle, True
        End If
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
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

Private Sub TempWbkHiddenRemove(ByVal wbk As Workbook)
    Dim app As Excel.Application
    
    Set app = wbk.Parent
    wbk.Close
    Set app = Nothing

End Sub

Public Sub UpdateByReplacingAllCodeLines(ByVal sync_vbc_source As VBComponent, _
                                         ByVal SYNC_TEST_WBK_SOURCE As Workbook, _
                                         ByVal SYNC_TEST_WBK_TARGET As Workbook)
' ----------------------------------------------------------------------------
' Synchronizes the code of a Sync-Target-Workbook's VBComponent by replacing
' the code lines with the code lines of the corrresponding Sync-Source-
' Workbook's VBComponent (sync_vbc_source).
' Note: Replacement by line is the only choice to replace the code of a Data
'       module.
' ----------------------------------------------------------------------------
    Const PROC = "UpdateByReplacingAllCodeLines"

    On Error GoTo eh
    Dim i           As Long
    Dim SourceComp  As clsComp
    Dim SourceCode  As Dictionary
    Dim v           As Variant
        
    '~~ Obtain the new code lines from the Sync-source-Workbook's component
    Set SourceComp = New clsComp
    With SourceComp
        Set .Wrkbk = SYNC_TEST_WBK_SOURCE
        Set .VBComp = sync_vbc_source
        Set SourceCode = .CodeLines
    End With
    
    With SYNC_TEST_WBK_TARGET.VBProject.VBComponents(sync_vbc_source.Name).CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines     ' Remove all lines from the cloned raw component
        For Each v In SourceCode                                    ' Insert the raw component's code lines
            i = i + 1
            .InsertLines i, SourceCode(v)
        Next v
    End With
                
xt: Set SourceComp = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function CollectOutdatedCommonComps(ByRef c_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the serviced Workbook's (mService.WbkServiced) used Common Components
' found still changed, i.e. not updated/renewed as Dictionary and additionally
' via (c_dct).
' ------------------------------------------------------------------------------
    Const PROC = "CollectOutdatedCommonComps"
    
    On Error GoTo eh
    Dim dctAll      As Dictionary
    Dim dctOutdated As New Dictionary
    Dim vbc         As VBComponent
    Dim fso         As New FileSystemObject
    Dim sOutdated   As String
    Dim v           As Variant
    Dim Comp        As clsComp
    Dim lAll        As Long
    Dim lRemaining  As Long
    Dim wbk         As Workbook
    Dim lUsed       As Long
    
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
                .CompName = vbc.Name
                Set .VBComp = vbc
                If .KindOfComp = enCommCompUsed Then
                    lUsed = lUsed + 1
                    If .Outdated Then
                        mService.AddAscByKey dctOutdated, .CompName, Comp
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
            mService.Progress(p_result:=dctOutdated.Count _
                            , p_of:=lUsed _
                            , p_op:="used Common Components are outdated" _
                            , p_comps:=sOutdated _
                            , p_dots:=lRemaining _
                             )
        Next v
    End With
    
    Set CollectOutdatedCommonComps = dctOutdated
    Set c_dct = dctOutdated
    mService.DsplyStatus _
    mService.Progress(p_result:=dctOutdated.Count _
                    , p_of:=lUsed _
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

