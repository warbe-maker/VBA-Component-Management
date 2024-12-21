Attribute VB_Name = "mUpdate"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mUpdate: Services for the updating Common Components, i.e.
' ======================== those, a corresponding Export-File exists in the
' Common-Components folder. The update is done either by re-importing the
' Export-File or by replacing the code by means of the mere code of an
' Export-File. However, the update
'
' Public services:
' ----------------
' ByCodeReplace Updates a VBComponent by replacing its code with the code
'               from another source (another VBComponent or an Export-File).
'               This service is the only choice for Data Modules, i.e. the
'               code of a Workbook or a Worksheet.
' ByReImport    Updates a VBComponent by removing, i.e. temporarily renaming
'               the target component) and re-importing an Export-File. This
'               service is the only choice for a UserForm.
'
' W. Rauschenberger, Berlin Aug. 2024
' ----------------------------------------------------------------------------
Private Const MONITORED_STEPS   As Long = 11
Private Const MONITOR_POS       As String = "20;400"

Private tMonitorFooter          As mMsg.udtMsgText
Private tMonitorStep            As mMsg.udtMsgText
Private sMonitorTitle           As String

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMe" & "." & sProc
End Function

Public Sub ByCodeReplace(ByVal b_source As clsComp, _
                         ByVal b_target As clsComp)
' ----------------------------------------------------------------------------
' Replaces the code in a component (b_target) with the code provided via
' a source component (b_source).
'
' Attention: This method is not applicable for UserForms since they have a
'            design component connected
' Note ....: Code replacement is the only update method applicable for Data
'            Modules but may as well be used for Standard- and Class-Modules.
' ----------------------------------------------------------------------------
    Const PROC = "ByCodeReplace"

    On Error GoTo eh
    Dim sComp As String
    
    If b_target.VBComp.Type = vbext_ct_MSForm _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "This service cannot be used to update a UserForm." & _
                                            "||A UserForm's update source can only be the two " & _
                                            "Export-Files used for re-importing also the UserForm's " & _
                                            "design (still tricky but the only choice)!"
    sComp = b_target.CodeCrrent.CompName
    b_target.CodeCrrent.ReplaceWith b_source.CodeCrrent
    
    With b_target
        Select Case CommonServiced.KindOfComponent(b_source.CompName)
            Case enCompCommonHosted:  Servicing.Log(sComp) = "Serviced Common Component  h o s t e d  updated by code replace"
            Case enCompCommonUsed:    Servicing.Log(sComp) = "Serviced Common Component  u s e d  updated by code replace"
        End Select
    End With
        
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ByReImport(ByVal b_comp_name As String, _
                      ByVal b_export_file As String, _
             Optional ByVal b_monitor As Boolean = False)
' ------------------------------------------------------------------------------
' Updates a component (b_comp_name) in a Workbook's (b_wbk) VBProject by
' re-importing an Export-File (b_export_file) with the following preconditions:
' - The Workbook (b_wbk) is open
' - The component (b_comp_name) exists in the VBProject of the Workbook
' - The source (b_export_file) is the full name of an Export-File
'
' W. Rauschenberger Berlin, Jan 2023
' ------------------------------------------------------------------------------
    Const PROC = "ByReImport"
    
    On Error GoTo eh
    Dim sTempName   As String
    Dim Comp        As clsComp
    Dim TmpWbk      As Workbook
    Dim lStep       As Long
    
    mBasic.BoP ErrSrc(PROC)
    MonitorInitiate "Update an outdated component", b_monitor
    
    With Serviced
        .Comp.CompName = b_comp_name
        
        '~~ Create and activate a hidden Workbook
        '~~ Because this may not be the first step the monitor initialization values are provided.
        '~~ They are ignored by the service when the monitor window is already initialized
        MonitorFirstStep m_step:=lStep _
                       , m_step_txt:="Activate a hidden temporary Workbook." _
                       , m_title:=sMonitorTitle _
                       , m_monitor:=b_monitor
        Set TmpWbk = TempWbkHidden()
        TmpWbk.Activate
                  
        sTempName = mComp.TempName(tn_wbk:=.Wrkbk, tn_vbc_name:=b_comp_name)
        
        MonitorStep lStep, "Rename the component '" & b_comp_name & "' to '" & sTempName & "'.", b_monitor
        '~~ Rename the already existing component
        .Wrkbk.VBProject.VBComponents(b_comp_name).Name = sTempName
        
        '~~ Outcomment the renamed component's code in order to avoid any compile conflict with
        '~~ the new imported component
        MonitorStep lStep, "Outcomment all code lines in the renamed component.", b_monitor
        OutCommentCode .Wrkbk, sTempName ' this had made it much less "reliablle"
        mBasic.TimedDoEvents ErrSrc(PROC)
        
        With .Wrkbk.VBProject
            '~~ Remove the renamed component.
            '~~ Note! The removal will in fact be done by the system when all runing procedured had finished.
            MonitorStep lStep, "Remove the renamed component.", b_monitor
            .VBComponents.Remove .VBComponents(sTempName)
                
            '~~ (Re-)import the public component's Export-File from the Common-Components folder
            MonitorStep lStep, "(Re-) import the Export File of the up-to-date version of the component.", b_monitor
            .VBComponents.Import b_export_file
                        
            '~~ Remove the activated hidden Workbook and re-activate the serviced Workbook
            MonitorStep lStep, "Remove the temporary activated and re-activate the serviced Workbook.", b_monitor
            TempWbkHiddenRemove TmpWbk
        End With
        .Wrkbk.Activate
        MonitorFooter "Successfully updated! (process monitor closes in 2 seconds)", b_monitor
        
        With .Comp
            .CompName = b_comp_name
            '~~ Copy the imported also as Export-File (do not export since this will crash)
            FSo.CopyFile b_export_file, .ExpFileFullName
        End With
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
        mMsg.Instance sMonitorTitle, True ' close any previous monitor window
        mMsg.Instance sMonitorTitle       ' establish a new monitor window
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
        mMsg.Instance sMonitorTitle, True
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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

