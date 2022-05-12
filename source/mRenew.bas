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
' performed to redisplay the still to be updated outdated components.
'
' Preconditions (not evaluated by the service):
' - r_wb_serviced is an open Workbook
' - r_comp_name is a components in its VB-Prtoject
' - r_exp_file_full-Name is a file in the "Common Components" folder
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------
    Const PROC = "Run"
    
    On Error GoTo eh
    Dim sTempName   As String
    Dim fso         As New FileSystemObject
    Dim Comp        As clsComp
    Dim TmpWbk      As Workbook
    Dim KoC         As Long
    
    mBasic.BoP ErrSrc(PROC)
    '~~ The service has been called directly for a single individual update
    Set Log = New clsLog
    Log.Service = SRVC_UPDATE_OUTDATED
    Set Comp = New clsComp
    With Comp
        Set .Wrkbk = r_wb_serviced
        .CompName = r_comp_name
        Log.ServicedItem = .VBComp
    End With
    EstablishTraceLogFile r_wb_serviced
    
    With r_wb_serviced.VBProject
        
        mTrc.BoC "Create and activate a hidden Workbook"
        Set TmpWbk = TempWbkHidden()
        TmpWbk.Activate
        mTrc.EoC "Create and activate a hidden Workbook"
        
        If mComp.Exists(wb:=r_wb_serviced, comp_name:=r_comp_name) Then
            '~~ Rename the already existing component ...
            mTrc.BoC "Rename the to-be-removed component"
            sTempName = mComp.TempName(tn_wb:=r_wb_serviced, tn_comp_name:=r_comp_name)
            .VBComponents(r_comp_name).Name = sTempName
            Log.Entry = "'" & r_comp_name & "' renamed to '" & sTempName & "'"
            mTrc.EoC "Rename the to-be-removed component"
            
            '~~ ... and outcomment its code.
            '~~ An (irregular) Workbook close may leave renamed components un-removed.
            '~~ When the Workbook is re-opened again any renamed component is deleted.
            '~~ However VB may fail when dedecting duplicate declarations before the
            '~~ renamed component can be deleted. Outcommenting the renamded component
            '~~ should prevent this.
            OutCommentCodeInRenamedComponent r_wb_serviced, sTempName ' this had made it much less "reliablele"
            mBasic.TimedDoEvents ErrSrc(PROC)
        
            '~~ Remove the renamed component (postponed thought)
            mTrc.BoC "Remove the renamde component"
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            Log.Entry = "'" & sTempName & "' removed (removal is postponed by Excel until process has finished)"
            mTrc.EoC "Remove the renamde component"
        End If
        
        '~~ (Re-)import the component
        mTrc.BoC "(Re-)Import the up-to-date Raw Common Component"
        .VBComponents.Import r_exp_file
        Log.Entry = "'" & r_comp_name & "' (re-)imported from '" & r_exp_file & "'"
        mTrc.EoC "(Re-)Import the up-to-date Raw Common Component"
                
        '~~ Export the re-newed Used Common Component
        mTrc.BoC "Export the Used Common Component"
        Set Comp = New clsComp
        With Comp
            Set Comp.Wrkbk = r_wb_serviced
            .CompName = r_comp_name
        End With
        .VBComponents(r_comp_name).Export Comp.ExpFileFullName
        Log.Entry = "'" & r_comp_name & "' exported to '" & Comp.ExpFileFullName & "'"
        mTrc.EoC "Export the Used Common Component"
        
        mTrc.BoC "Remove the activated hidden Workbook and re-activate the processed Workbook"
        TempWbkHiddenRemove TmpWbk
        r_wb_serviced.Activate
        mTrc.EoC "Remove the activated hidden Workbook and re-activate the processed Workbook"
  
        '~~ Update the used Common Component's RevisionNumber now equal to the RevisionNumber of the re-imported component's Export file
        With Comp
            Set Comp.Wrkbk = r_wb_serviced
            .CompName = r_comp_name
            KoC = .KindOfComp ' this establishes the .Raw instance
            Log.Entry = "Used Common Component renewed/updated by (re-)import of the Raw's Export File (" & .Raw.SavedExpFileFullName & ")"
            .RevisionNumber = .Raw.RevisionNumber
            .DueModificationWarning = False
        End With
        '~~ The re-call of the update service with the option uo_modeless=True ends with the
        '~~ display of a new update message in case there are still outdated components to be
        '~~ updated.
        mCompMan.UpdateOutdatedCommonComponents uo_wb:=r_wb_serviced, uo_hosted:=r_hosted, uo_modeless:=True
    
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
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
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
