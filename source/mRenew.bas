Attribute VB_Name = "mRenew"
Option Explicit

Public Sub ByImport(ByRef bi_wbk_serviced As Workbook, _
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
        Log.Service = mCompManClient.SRVC_UPDATE_OUTDATED
        Set Comp = New clsComp
        With Comp
            Set .Wrkbk = bi_wbk_serviced
            .CompName = bi_vbc_name
            Log.ServicedItem = .VBComp
        End With
        EstablishExecTraceFile bi_wbk_serviced
    End If
    
    With bi_wbk_serviced.VBProject
        
        mTrc.BoC "Create and activate a hidden Workbook"
        Set TmpWbk = TempWbkHidden()
        TmpWbk.Activate
        mTrc.EoC "Create and activate a hidden Workbook"
        
        If mComp.Exists(xst_wbk:=bi_wbk_serviced, xst_vbc_name:=bi_vbc_name) Then
            '~~ Rename the already existing component ...
            mTrc.BoC "Rename the to-be-removed component"
            sTempName = mComp.TempName(tn_wbk:=bi_wbk_serviced, tn_vbc_name:=bi_vbc_name)
            .VBComponents(bi_vbc_name).name = sTempName
            mTrc.EoC "Rename the to-be-removed component"
            
            '~~ ... and outcomment its code.
            '~~ An (irregular) Workbook close may leave renamed components un-removed.
            '~~ When the Workbook is re-opened again any renamed component is deleted.
            '~~ However VB may fail when dedecting duplicate declarations before the
            '~~ renamed component can be deleted. Outcommenting the renamded component
            '~~ should prevent this.
            mOutdated.OutCommentCodeInRenamedComponent bi_wbk_serviced, sTempName ' this had made it much less "reliablele"
            mBasic.TimedDoEvents ErrSrc(PROC)
        
            '~~ Remove the renamed component (postponed thought)
            mTrc.BoC "Remove the renamde component"
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            mTrc.EoC "Remove the renamde component"
        End If
        
        '~~ (Re-)import the component
        mTrc.BoC "(Re-)Import the up-to-date Raw Common Component"
        .VBComponents.Import bi_exp_file
        Log.Entry = "Renewed/updated by (re-)import of '" & bi_exp_file & "'"
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

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mRenew." & s
End Function

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
