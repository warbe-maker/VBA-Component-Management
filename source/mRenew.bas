Attribute VB_Name = "mRenew"
Option Explicit

Public Sub ByImport(ByRef bi_wb_serviced As Workbook, _
                    ByVal bi_comp_name As String, _
                    ByVal bi_exp_file As String)
' ------------------------------------------------------------------------------
' This service renews a Workbook's (bi_wb_serviced) component (bi_comp_name) by
' re-importing a more up-to-date component's export file (bi_exp_file). When
' the service is performed directly by Application.Run (bi_via_app_run = True)
' it calls mCompMan.Update again to redisplay the still to be updated outdated
' components.
' Preconditions (not evaluated by the service):
' - bi_wb_serviced is an open Workbook
' - bi_comp_name is a components in its VB-Prtoject
' - bi_exp_file_full-Name is a file in the "Common Components" folder
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------
    Const PROC = "ByImport"
    
    On Error GoTo eh
    Dim sTempName   As String
    Dim fso         As New FileSystemObject
    Dim Comp        As clsComp
    Dim TmpWbk      As Workbook
    Dim KoC         As Long
    
    mBasic.BoP ErrSrc(PROC)
    If Log Is Nothing Then
        '~~ The service has been called directly for a single individual update
        Set Log = New clsLog
        Log.Service = SRVC_UPDATE_OUTDATED
        Set Comp = New clsComp
        With Comp
            Set .Wrkbk = bi_wb_serviced
            .CompName = bi_comp_name
            Log.ServicedItem = .VBComp
        End With
        EstablishTraceLogFile bi_wb_serviced
    End If
    
    With bi_wb_serviced.VBProject
        
        mTrc.BoC "Create and activate a hidden Workbook"
        Set TmpWbk = TempWbkHidden()
        TmpWbk.Activate
        mTrc.EoC "Create and activate a hidden Workbook"
        
        If mComp.Exists(wb:=bi_wb_serviced, comp_name:=bi_comp_name) Then
            '~~ Rename the already existing component ...
            mTrc.BoC "Rename the to-be-removed component"
            sTempName = mComp.TempName(tn_wb:=bi_wb_serviced, tn_comp_name:=bi_comp_name)
            .VBComponents(bi_comp_name).Name = sTempName
            Log.Entry = "'" & bi_comp_name & "' renamed to '" & sTempName & "'"
            mTrc.EoC "Rename the to-be-removed component"
            
            '~~ ... and outcomment its code.
            '~~ An (irregular) Workbook close may leave renamed components un-removed.
            '~~ When the Workbook is re-opened again any renamed component is deleted.
            '~~ However VB may fail when dedecting duplicate declarations before the
            '~~ renamed component can be deleted. Outcommenting the renamded component
            '~~ should prevent this.
            OutCommentCodeInRenamedComponent bi_wb_serviced, sTempName ' this had made it much less "reliablele"
            mBasic.TimedDoEvents ErrSrc(PROC)
        
            '~~ Remove the renamed component (postponed thought)
            mTrc.BoC "Remove the renamde component"
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            Log.Entry = "'" & sTempName & "' removed (removal is postponed by Excel until process has finished)"
            mTrc.EoC "Remove the renamde component"
        End If
        
        '~~ (Re-)import the component
        mTrc.BoC "(Re-)Import the up-to-date Raw Common Component"
        .VBComponents.Import bi_exp_file
        Log.Entry = "'" & bi_comp_name & "' (re-)imported from '" & bi_exp_file & "'"
        mTrc.EoC "(Re-)Import the up-to-date Raw Common Component"
                
        '~~ Export the re-newed Used Common Component
        mTrc.BoC "Export the Used Common Component"
        Set Comp = New clsComp
        With Comp
            Set Comp.Wrkbk = bi_wb_serviced
            .CompName = bi_comp_name
        End With
        .VBComponents(bi_comp_name).Export Comp.ExpFileFullName
        Log.Entry = "'" & bi_comp_name & "' exported to '" & Comp.ExpFileFullName & "'"
        mTrc.EoC "Export the Used Common Component"
        
        mTrc.BoC "Remove the activated hidden Workbook and re-activate the processed Workbook"
        TempWbkHiddenRemove TmpWbk
        bi_wb_serviced.Activate
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

Public Sub Run(ByRef r_wb_serviced As Workbook, _
               ByVal r_comp_name As String, _
               ByVal r_exp_file As String, _
      Optional ByVal r_hosted As String = vbNullString)
' ------------------------------------------------------------------------------
' Renews a Workbook's (r_wb_serviced) component (r_comp_name) by re-importing a
' more up-to-date component's export file (r_exp_file). The service is dedicated
' to be performed via Application.Run. After execution mCompMan.Update... is
' performed to redisplay the outstanding to be updated outdated components.
'
' Preconditions (not evaluated by the service):
' - r_wb_serviced is an open Workbook
' - r_comp_name is a components in its VB-Prtoject
' - r_exp_file_full-Name is a file in the "Common Components" folder
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------
    Const PROC              As String = "Run"
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
    MonitorFooter.FontColor = rgbBlue
    MonitorFooter.FontSize = 9
    mMsg.MsgInstance MonitorTitle, True ' close any previous monitor window
    mMsg.MsgInstance MonitorTitle       ' establish a new monitor window
    MonitorStep.MonoSpaced = True
    MonitorStep.FontSize = 9
    Set Log = New clsLog
    Log.Service(True) = SRVC_UPDATE_OUTDATED ' new log file when older 1 day
    
    '~~ Save the serviced Workbook when yet not saved (initialize the monitor window)
    If Not r_wb_serviced.Saved Then
        Step = Step + 1
        MonitorStep.Text = Step & ". Save serviced Workbook '" & r_wb_serviced.Name & "'"
        mMsg.Monitor mon_title:=MonitorTitle _
                   , mon_text:=MonitorStep _
                   , mon_steps_displayed:=MONITORED_STEPS _
                   , mon_width_min:=70 _
                   , mon_pos:=MONITOR_POS
        mService.SaveWbk r_wb_serviced
    End If
    
    Set Comp = New clsComp
    With Comp
        Set .Wrkbk = r_wb_serviced
        .CompName = r_comp_name
        Log.ServicedItem = .VBComp
    End With
    EstablishTraceLogFile r_wb_serviced
    
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
        
        
    With r_wb_serviced.VBProject
        If mComp.Exists(wb:=r_wb_serviced, comp_name:=r_comp_name) Then
        
            '~~ Rename an already existing component
            sTempName = mComp.TempName(tn_wb:=r_wb_serviced, tn_comp_name:=r_comp_name)
            Step = Step + 1
            MonitorStep.Text = Step & ". Rename the component '" & r_comp_name & "' to '" & sTempName & "'."
            mMsg.Monitor MonitorTitle, MonitorStep
            .VBComponents(r_comp_name).Name = sTempName
            
            '~~ Outcomment the renamed component's code
            Step = Step + 1
            MonitorStep.Text = Step & ". Outcomment all code lines in the renamed component."
            mMsg.Monitor MonitorTitle, MonitorStep
            OutCommentCodeInRenamedComponent r_wb_serviced, sTempName ' this had made it much less "reliablele"
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
        .VBComponents.Import r_exp_file
        Log.Entry = "'" & r_comp_name & "' (re-)imported from '" & r_exp_file & "'"
                
        '~~ Export the re-newed Used Common Component
        Step = Step + 1
        MonitorStep.Text = Step & ". Export the (re-)imported component."
        mMsg.Monitor MonitorTitle, MonitorStep
        Set Comp = New clsComp
        With Comp
            Set Comp.Wrkbk = r_wb_serviced
            .CompName = r_comp_name
        End With
        .VBComponents(r_comp_name).Export Comp.ExpFileFullName
        
        '~~ Remove the activated hidden Workbook and re-activate the serviced Workbook
        Step = Step + 1
        MonitorStep.Text = Step & ". Remove the temporary activated and re-activate the serviced Workbook."
        mMsg.Monitor MonitorTitle, MonitorStep
        TempWbkHiddenRemove TmpWbk
        r_wb_serviced.Activate
  
        '~~ Set the updated component's RevisionNumber = the RevisionNumber of the re-imported component's Export file
        Step = Step + 1
        MonitorStep.Text = Step & ". Update the updated component's revision number."
        mMsg.Monitor MonitorTitle, MonitorStep
        With Comp
            Set Comp.Wrkbk = r_wb_serviced
            .CompName = r_comp_name
            KoC = .KindOfComp ' this establishes the .Raw instance
            Log.Entry = "Used Common Component renewed/updated by (re-)import of the Raw's Export File (" & .Raw.SavedExpFileFullName & ")"
            .RevisionNumber = .Raw.RevisionNumber
            .DueModificationWarning = False
        End With

        '~~ Re-call the update service in order to display still outdated components if any
        MonitorFooter.Text = "Outdated component '" & r_comp_name & "' successfully updated."
        mMsg.MonitorFooter MonitorTitle, MonitorFooter
        mCompMan.UpdateOutdatedCommonComponents uo_wb_serviced:=r_wb_serviced, uo_hosted:=r_hosted
    
    End With
    
xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mRenew." & s
End Function

Private Sub OutCommentCodeInRenamedComponent(ByRef oc_wb As Workbook, _
                                             ByVal oc_component As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutCommentCodeInRenamedComponent"
    
    On Error GoTo eh
    Dim i As Long
    
    mBasic.BoP ErrSrc(PROC)
    With oc_wb.VBProject.VBComponents(oc_component).CodeModule
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

Public Function Outdated(ByVal od_wb_serviced As Workbook) As Dictionary
' ------------------------------------------------------------------------------
' Returns the serviced Workbook's (od_wb_serviced) outdated Common Components as
' Dictionary.
' ------------------------------------------------------------------------------
    Const PROC = "Outdated"
    
    On Error GoTo eh
    Dim dctAll          As Dictionary
    Dim dctOutdated     As New Dictionary
    Dim vbc             As VBComponent
    Dim fso             As New FileSystemObject
    Dim sOutdated        As String
    Dim v               As Variant
    Dim Comp            As clsComp
    Dim Comps           As New clsComps
    Dim lAll            As Long
    Dim lRemaining      As Long
    
    mBasic.BoP ErrSrc(PROC)
    Set dctAll = mService.AllComps(od_wb_serviced)
    SaveWbk od_wb_serviced

    With od_wb_serviced.VBProject
        lAll = .VBComponents.Count
        lRemaining = lAll
        For Each v In dctAll
            Set vbc = dctAll(v)
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = od_wb_serviced
                .CompName = vbc.Name
                Set .VBComp = vbc
                If .KindOfComp = enCommCompUsed Then
                    If .Outdated Then
                        AddAscByKey dctOutdated, .CompName, Comp
                        sOutdated = .CompName
                    End If ' .Outdated
                End If ' Used Common Component
            End With
            Set Comp = Nothing
            lRemaining = lRemaining - 1
            Application.StatusBar = _
            mService.Progress(p_service:=Log.Service _
                            , p_result:=dctOutdated.Count _
                            , p_of:=lAll _
                            , p_op:="used Common Components are outdated" _
                            , p_comps:=sOutdated _
                            , p_dots:=lRemaining _
                             )
        Next v
    End With
    
    Set Outdated = dctOutdated
    Application.StatusBar = vbNullString
    Application.StatusBar = _
    mService.Progress(p_service:=Log.Service _
                    , p_result:=dctOutdated.Count _
                    , p_of:=lAll _
                    , p_op:="used Common Components are outdated" _
                    , p_comps:=sOutdated _
                     )
    
xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function



Private Function TempWbkHidden() As Workbook
    Dim app As Excel.Application
    
    Set app = CreateObject("Excel.Application")
    app.Visible = False
    Set TempWbkHidden = app.Workbooks.Add()

End Function

Private Sub TempWbkHiddenRemove(ByVal wb As Workbook)
    Dim app As Excel.Application
    
    Set app = wb.Parent
    wb.Close
    Set app = Nothing

End Sub
