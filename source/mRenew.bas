Attribute VB_Name = "mRenew"
Option Explicit

Public Sub ByImport( _
              ByRef rn_wb As Workbook, _
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
    Dim sTempName       As String
    Dim fso             As New FileSystemObject
    Dim Comp           As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    SaveWbk rn_wb
    
    With rn_wb.VBProject
        
        '~~ Find a free/unused temporary name and re-name the outdated component
        mTrc.BoC "Rename the to-be-re-imported component"
        If mComp.Exists(wb:=rn_wb, comp_name:=rn_comp_name) Then
            sTempName = mComp.TempName(tn_wb:=rn_wb, tn_comp_name:=rn_comp_name)
            '~~ Rename the component when it already exists
            .VBComponents(rn_comp_name).Name = sTempName
            Log.Entry = "'" & rn_comp_name & "' renamed to '" & sTempName & "'"
            .VBComponents.Remove .VBComponents(sTempName) ' will not take place until process has ended!
            Log.Entry = "'" & sTempName & "' removed (removal is postponed by Excel until process has finished)"
        End If
        mTrc.EoC "Rename the to-be-re-imported component"
        
        '~~ (Re-)import the component
        mTrc.BoC "Import the Raw Common Component"
        .VBComponents.Import rn_raw_exp_file_full_name
        Log.Entry = "'" & rn_comp_name & "' (re-)imported from '" & rn_raw_exp_file_full_name & "'"
        mTrc.EoC "Import the Raw Common Component"
        
        '~~ Export the re-newd Used Common Component
        mTrc.BoC "Export the Used Common Component"
        Set Comp = New clsComp
        With Comp
            Set Comp.Wrkbk = rn_wb
            .CompName = rn_comp_name
        End With
        .VBComponents(rn_comp_name).Export Comp.ExpFileFullName
        Log.Entry = "'" & rn_comp_name & "' exported to '" & Comp.ExpFileFullName & "'"
        mTrc.EoC "Export the Used Common Component"
        
        '~~ When Excel closes the Workbook with the subsequent Workbook save it may be re-opened
        '~~ and the update process will continue with the next outdated Used Common Component.
        '~~ The (irregular) Workbook close however may leave the renamed components un-removed.
        '~~ When the Workbook is opened again these renamed component may cause duplicate declarations.
        '~~ To prevent this the code in the renamed component is dleted.
        ' EliminateCodeInRenamedComponent sTempName ' this had made it much less "reliablele"
        
        SaveWbk rn_wb ' This "crahes" every now an then though I've tried a lot
    
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

Private Sub EliminateCodeInRenamedComponent(ByVal temp_comp_name As String, _
                                            ByVal temp_un_comment_only As Boolean)
' ----------------------------------------------------------------------------
' Removes all code lines in the renamed component named (temp_comp_name).
' Background:
' When Excel closes the Workbook with the subsequent Workbook save it may be
' re-opened and the update process will continue with the next outdated Used
' Common Component. The (irregular) Workbook close however may leave the
' renamed components un-removed. When the Workbook is opened again these
' renamed components may cause duplicate declarations. To prevent this the
' code in the renamed component is deleted.
' ----------------------------------------------------------------------------
    Const PROC = "EliminateCodeInRenamedComponent"
    
    On Error GoTo eh
    Dim i As Long
    
    mBasic.BoP ErrSrc(PROC)
    With mService.Serviced.VBProject.VBComponents(temp_comp_name)
        With .CodeModule
            If Not temp_un_comment_only Then
                .DeleteLines 1, .CountOfLines
            Else
                For Each v In .Lines
                    v
                Next v
            End If
        End With
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SaveWbk(ByRef rs_wb As Workbook)
    Const PROC = "SaveWbk"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    
    Application.EnableEvents = False
    '~~ This is the action where the update process may lead to the effect that Excel closes the Workbook
    '~~ without having deleted the renamed components!
    Log.Entry = "Workbook about to be saved - with the outdate Used Common Component already updated"
    Log.Entry = "DoEvents delayed continuation for " & TimedDoEvents & " msec)"
    rs_wb.Save
    Log.Entry = "Workbook saved (DoEvents delayed continuation for " & TimedDoEvents & " msec)"
    Application.EnableEvents = True

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function TimedDoEvents() As String
    mBasic.TimerBegin
    DoEvents
    TimedDoEvents = Format(mBasic.TimerEnd, "00000")
End Function

