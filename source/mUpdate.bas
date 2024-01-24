Attribute VB_Name = "mUpdate"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mUpdate: Services for the update of a VBComponent, either
' ------------------------ by re-importing an Export-File or by replacing the
'                          code.
' Public services:
' - ByCodeReplace Updates a VBComponent by replacing its code with the code
'                 from another VBComponent.
' - ByReImport    Updates a VBComponent by removing and re-importing an
'                 Export-File.
'
' ----------------------------------------------------------------------------
Private Const MONITORED_STEPS   As Long = 11
Private Const MONITOR_POS       As String = "20;400"

Private tMonitorFooter          As mMsg.udtMsgText
Private tMonitorStep            As mMsg.udtMsgText
Private sMonitorTitle           As String

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMe" & "." & sProc
End Function

Public Sub ByCodeReplace(ByVal b_source_comp As clsComp, _
                         ByVal b_target_comp As clsComp)
' ----------------------------------------------------------------------------
' Replaces the code in a target component's (b_target_comp) CodeModule with
' the code provided through a source component (b_source_comp).
' Attention: This method is not applicable for UserForms since they have a
'            design component connected
' Note ....: Code replacement is the only update method applicable for Data
'            Modules but may as well be used for Standard- and Class-Modules.
' ----------------------------------------------------------------------------
    Const PROC = "ByCodeReplace"

    On Error GoTo eh
    Dim CodeTarget  As New clsCode
    Dim CodeSource  As New clsCode
    
    CodeSource.Source = b_source_comp.VBComp.CodeModule
    CodeTarget.ReplaceWith r_target_vbcm:=b_target_comp.VBComp.CodeModule _
                         , r_source_code:=CodeSource
    
    With b_target_comp
        Select Case .KindOfComp
            Case enCommCompHosted
                Services.ServicedItemLogEntry "Outdated Common Component hosted updated by code replace"
            Case enCommCompUsed
                Services.ServicedItemLogEntry "Outdated Common Component used updated by code replace"
        End Select
    End With
        
xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ByReImport(ByVal b_wbk_target As Workbook, _
                      ByVal b_vbc_name As String, _
                      ByVal b_exp_file As String, _
             Optional ByVal b_monitor As Boolean = False)
' ------------------------------------------------------------------------------
' Updates a target component of which the source has changed by re-importing the
' source component's export-file to the target VB-Project.
' Preconditions:
' - The source Workbook is open
' - The VBComponent (b_vbc_name) is a components in the source Workbook's VB-Project
' - The Export-file (b_exp_file is from the source VBComponent
'
' W. Rauschenberger Berlin, Jan 2023
' ------------------------------------------------------------------------------
    Const PROC              As String = "ByReImport"
    
    On Error GoTo eh
    Dim sTempName       As String
    Dim Comp            As clsComp
    Dim TmpWbk          As Workbook
    Dim lStep           As Long
    Dim sLastModAtDateTime          As String
    
    mBasic.BoP ErrSrc(PROC)
    MonitorInitiate "Update an outdated component", b_monitor
    
    Set Comp = New clsComp
    With Comp
        .Wrkbk = b_wbk_target
        .CompName = b_vbc_name
    End With
    
    '~~ Create and activate a hidden Workbook
    '~~ Because this may not be the first step the monitor initialization values are provided.
    '~~ They are ignored by the service when the monitor window is already initialized
    MonitorFirstStep m_step:=lStep _
                   , m_step_txt:="Activate a hidden temporary Workbook." _
                   , m_title:=sMonitorTitle _
                   , m_monitor:=b_monitor
    Set TmpWbk = TempWbkHidden()
    TmpWbk.Activate
              
    With b_wbk_target.VBProject
        '~~ Rename an already existing component
        sTempName = mComp.TempName(tn_wbk:=b_wbk_target, tn_vbc_name:=b_vbc_name)
        MonitorStep lStep, "Rename the component '" & b_vbc_name & "' to '" & sTempName & "'.", b_monitor
        .VBComponents(b_vbc_name).Name = sTempName
        
        '~~ Outcomment the renamed component's code
        MonitorStep lStep, "Outcomment all code lines in the renamed component.", b_monitor
        OutCommentCode b_wbk_target, sTempName ' this had made it much less "reliablele"
        mBasic.TimedDoEvents ErrSrc(PROC)
        
        '~~ Remove the renamed component (postponed thought)
        MonitorStep lStep, "Remove the renamed component.", b_monitor
        .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
                
        '~~ (Re-)import the component
        MonitorStep lStep, "(Re-) import the Export File of the up-to-date version of the component.", b_monitor
        .VBComponents.Import b_exp_file
                        
        '~~ Remove the activated hidden Workbook and re-activate the serviced Workbook
        MonitorStep lStep, "Remove the temporary activated and re-activate the serviced Workbook.", b_monitor
        TempWbkHiddenRemove TmpWbk
        b_wbk_target.Activate
          
        MonitorFooter "Successfully updated! (process monitor closes in 2 seconds)", b_monitor
    End With
    
    Set Comp = Nothing
    Set Comp = New clsComp
    With Comp
        .Wrkbk = b_wbk_target
        .CompName = b_vbc_name
        CommComps.CompName = b_vbc_name
        '~~ Export the re-imported component to ensure an up-to-date Export-File
        fso.CopyFile b_exp_file, .ExpFileFullName, True
        '~~ Update the Revision-Number
        sLastModAtDateTime = CommComps.LastModAtDateTime
        Select Case .KindOfComp
            Case enCommCompUsed, enCommCompHosted
                If .LastModAtDateTime <> sLastModAtDateTime Then
                    .LastModAtDateTime = sLastModAtDateTime
                End If
        End Select
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub MonitorInitiate(ByVal m_title As String, _
                   Optional ByVal m_monitor As Boolean = False)
' ------------------------------------------------------------------------------
' Initiates a monitored process
' ------------------------------------------------------------------------------
    If m_monitor Then
        sMonitorTitle = m_title
        tMonitorFooter.FontColor = rgbDarkGreen
        tMonitorFooter.FontSize = 9
        mMsg.MsgInstance sMonitorTitle, True ' close any previous monitor window
        mMsg.MsgInstance sMonitorTitle       ' establish a new monitor window
        tMonitorStep.MonoSpaced = True
        tMonitorStep.FontSize = 9
    End If

End Sub

Private Sub MonitorFooter(ByVal m_footer As String, _
                 Optional ByVal m_monitor As Boolean = False)
' ------------------------------------------------------------------------------
' Displays a monitored processes footer
' ------------------------------------------------------------------------------
    
    If m_monitor Then
        tMonitorFooter.Text = m_footer
        mMsg.MonitorFooter sMonitorTitle, tMonitorFooter
        Application.Wait Now() + TimeValue("0:00:02")
        mMsg.MsgInstance sMonitorTitle, True
    End If

End Sub

Private Sub MonitorStep(ByRef m_step As Long, _
                    ByVal m_step_txt As String, _
           Optional ByVal m_monitor As Boolean = False)
' ------------------------------------------------------------------------------
' Displays a monitored step
' ------------------------------------------------------------------------------
    
    If m_monitor Then
        m_step = m_step + 1
        tMonitorStep.Text = m_step & ". " & m_step_txt
        mMsg.Monitor sMonitorTitle, tMonitorStep
    End If

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

Private Sub OutCommentCode(ByVal o_wbk As Workbook, _
                                             ByVal o_component As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "OutCommentCode"
    
    On Error GoTo eh
    Dim i As Long
    
    mBasic.BoP ErrSrc(PROC)
    With o_wbk.VBProject.VBComponents(o_component).CodeModule
        For i = 1 To .CountOfLines
            .ReplaceLine i, "'" & .Lines(i, 1)
        Next i
    End With

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub MonitorFirstStep(ByRef m_step As Long, _
                             ByVal m_step_txt As String, _
                             ByVal m_title As String, _
                    Optional ByVal m_monitor As Boolean = False)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    If m_monitor Then
        m_step = m_step + 1
        tMonitorStep.Text = m_step & ". " & m_step_txt
        mMsg.Monitor mon_title:=m_title _
                   , mon_text:=tMonitorStep _
                   , mon_steps_displayed:=MONITORED_STEPS _
                   , mon_width_min:=70 _
                   , mon_pos:=MONITOR_POS
    End If

End Sub

