Attribute VB_Name = "mRenew"
Option Explicit

Public Sub ByImport(ByRef rn_wb As Workbook, _
                    ByVal rn_comp_name As String, _
                    ByVal rn_raw_exp_file_full_name As String)
' ----------------------------------------------------------------------------
' Renews/replaces the component (rn_comp_name) in Workbook (rn_wb) by
' importing the Export-File (rn_raw_exp_file_full_name).
' Note: Because a module cannot be deleted it is renamed and deleted. The
'       rename puts it out of the way, deletion is done by the system when the
'       process has ended.
' ----------------------------------------------------------------------------
    Const PROC = "ByImport"
    
    On Error GoTo eh
    Dim sTempName   As String
    Dim fso         As New FileSystemObject
    Dim Comp        As clsComp
    Dim TmpWbk      As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    With rn_wb.VBProject
        
        mTrc.BoC "Create and activate a hidden Workbook"
        Set TmpWbk = TempWbkHidden()
        TmpWbk.Activate
        mTrc.EoC "Create and activate a hidden Workbook"
        
        If mComp.Exists(wb:=rn_wb, comp_name:=rn_comp_name) Then
            '~~ Rename the component when it already exists ...
            mTrc.BoC "Rename the to-be-removed component"
            sTempName = mComp.TempName(tn_wb:=rn_wb, tn_comp_name:=rn_comp_name)
            .VBComponents(rn_comp_name).Name = sTempName
            Log.Entry = "'" & rn_comp_name & "' renamed to '" & sTempName & "'"
            mTrc.EoC "Rename the to-be-removed component"
            
            '~~ ... and outcomment its code.
            '~~ An (irregular) Workbook close may leave renamed components un-removed.
            '~~ When the Workbook is re-opened again any renamed component is deleted.
            '~~ However VB may fail when dedecting duplicate declarations before the
            '~~ renamed component can be deleted. Outcommenting the renamded component
            '~~ should prevent this.
            OutCommentCodeInRenamedComponent rn_wb, sTempName ' this had made it much less "reliablele"
            Debug.Print "DoEvents paused the execution for " & mService.TimedDoEvents & " msec"
        
            '~~ Remove the renamed component (postponed thought)
            mTrc.BoC "Remove the renamde component"
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            Log.Entry = "'" & sTempName & "' removed (removal is postponed by Excel until process has finished)"
            mTrc.EoC "Remove the renamde component"
        End If
        
        '~~ (Re-)import the component
        mTrc.BoC "(Re-)Import the up-to-date Raw Common Component"
        .VBComponents.Import rn_raw_exp_file_full_name
        Log.Entry = "'" & rn_comp_name & "' (re-)imported from '" & rn_raw_exp_file_full_name & "'"
        mTrc.EoC "(Re-)Import the up-to-date Raw Common Component"
                
        '~~ Export the re-newed Used Common Component
        mTrc.BoC "Export the Used Common Component"
        Set Comp = New clsComp
        With Comp
            Set Comp.Wrkbk = rn_wb
            .CompName = rn_comp_name
        End With
        .VBComponents(rn_comp_name).Export Comp.ExpFileFullName
        Log.Entry = "'" & rn_comp_name & "' exported to '" & Comp.ExpFileFullName & "'"
        mTrc.EoC "Export the Used Common Component"
        
        mTrc.BoC "Remove the activated hidden Workbook and re-activate the processed Workbook"
        TempWbkHiddenRemove TmpWbk
        rn_wb.Activate
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
