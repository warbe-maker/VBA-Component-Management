Attribute VB_Name = "mExport"
Option Explicit
' ----------------------------------------------------------------------------
' Standard-Module mExport: Services specifically for the export of changed or
' ======================== all components.
'
' Public services:
' ----------------
' All               Exports all VBComponentnts whether the code has changed
'                   or not
' ChangedComponents Exports all VBComponents of which the code has changed,
'                   i.e. a temporary Export-File differs from the
'                       regular Export-File (of the previous code change).
' ExpFileFolderPath Returns a serviced Workbook's path for all Export-Files
'                   whereby the name of the folder is the current configured
'                   one (fefaulting  to 'source'). When no Export-Folder
'                   exists, one is created. In case an outdated export folder
'                   exists, i.e. one with an outdated name, this one is
'                   renamed instead.
' ----------------------------------------------------------------------------

Public Sub ChangedComponents(Optional ByVal c_comp As String = vbNullString)
' ----------------------------------------------------------------------------
' - Exports all (see exception) components the code had been modified
' - Registers a due warning message when a Used Common Component had been
'   modified in the Workbook which uses but not hosts it.
' Note: When a single component is provided (c_comp) only this component is
'       exported. This option is only used during testings.
' ----------------------------------------------------------------------------
    Const PROC = "ChangedComponents"
    
    On Error GoTo eh
    Dim Comp                    As clsComp
    Dim sComp                   As String
    Dim wbk                     As Workbook
    Dim vbc                     As VBComponent
    Dim bChanged                As Boolean
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Prevent any action when the required preconditins are not met
    If Servicing.Denied(mCompManClient.SRVC_EXPORT_CHANGED) Then GoTo xt
    
    Set wbk = Servicing.ServicedWbk
    With Prgrss
        .Operation = "exported"
        .ItemsTotal = Serviced.Wrkbk.VBProject.VBComponents.Count
        .Figures = True
        .DoneItemsInfo = True
    End With
    
    With Serviced
        For Each vbc In .Wrkbk.VBProject.VBComponents
            sComp = vbc.Name
            If c_comp <> vbNullString And sComp = c_comp _
            Or c_comp = vbNullString Then
                .Comp.CompName = vbc.Name
                With .Comp
                    bChanged = .Changed
                    
                    Select Case True
                        ' =====================================
                        '                         --- Rules ---
                        '                         1 2 3 4 5 6 7
                        ' ----------------------- -------------
                        ' Conditions:
                        ' C1 Changed  *)          N Y Y Y Y Y Y
                        ' C2 Common Component     - N Y Y Y Y Y
                        ' C3 Export exists        - - N
                        ' C3 Up-to-date **)       - - Y - - - -
                        ' C4 Public               - - N Y Y Y Y
                        ' C5 Pending              - - N Y Y N Y
                        ' C6 Conflict             - - N N Y
                        ' ----------------------- -------------
                        ' Actions:
                        ' A1 None, skip           x
                        ' A2 Export                 x x x
                        ' A3 Update CompMan dat       x x
                        ' A4 Register Pending           x
                        ' A5 Update Pending.dat         x
                        ' A6 Resolve Conflict ***)         x
                        ' ==================================
                        '   *) Diff or no exp. file
                        '  **) Current code is up-to-date or last export indicates an up-to-date modification base
                        ' ***) Backout modification
                        
                        Case Not bChanged And .IsCommon
                            '~~ Common Component which not has changed
                            Prgrss.ItemSkipped
        
                        Case Not bChanged
                            '~~ Any other component which not has changed
                            Prgrss.ItemSkipped
                        
                        Case Not .IsCommon
                            '~~ Any other but a Common Component the code has changed
                            .Export
                            Servicing.Log(sComp) = "Modified VBComponent e x p o r t e d !"
                            Prgrss.ItemDone = sComp
                        
                        Case Not .IsCommCompPublic And Not .IsCommCompPending
                            '~~ Common Component the code has changed, yet no public version, not yet pending
                            .Export
                            .SetServicedProperties
                            .CodeExprtd.Source = .ExpFileFullName
                            CommonPending.Register Comp
                            With Servicing
                                .Log(sComp) = "Serviced Common Component modified: E x p o r t e d !"
                                .Log(sComp) = "Serviced Common Component modified: Registered pending release"
                            End With
                            Prgrss.ItemDone = sComp
                            
                        Case .IsCommCompPublic And Not .IsCommCompPending
                            '~~ Common Component the code has changed, already public, yet not pending release
                            '~~ (may as well be up-to-date but never exported)
                            .Export
                            .SetServicedProperties
                            If Not .IsCommCompUpToDate Then
                                '~~ This indicates that the exported Common Component is not the result of
                                '~~ a manually imported public Common Component's Export File
                                CommonPending.Register Serviced.Comp
                                .CodePnding.Source = CommonPending.LastModExpFile(sComp)
                                With Servicing
                                    .Log(sComp) = "Serviced Common Component modified: E x p o r t e d !"
                                    .Log(sComp) = "Serviced Common Component modified: Properties in CommComps.dat updated"
                                    .Log(sComp) = "Serviced Common Component modified: Registered pending release"
                                End With
                            End If
                            Prgrss.ItemDone = sComp
                        
                        Case .IsCommCompPublic And Not .IsCommCompPending And Not .IsCommCompUpToDate
                            '~~ When the Common Component is not up-to-date it must not be changed/modified!
                            '~~ Choices are:
                            ' ~~ - When there's yet no public version, export without registration pending release
                            '~~  - When there's a public version, discarding any changes by re-importing the public version.
                            ResolveExportConflict Comp
                            Prgrss.ItemDone = sComp
                        
                        Case .IsCommCompPublic And Not CommonPending.Conflicts(Serviced.Comp)
                            '~~ Changed, common public component, pending, not conflicting
                            .Export
                            .SetServicedProperties
                            CommonPending.Register Serviced.Comp
                            .CodePnding.Source = CommonPending.LastModExpFile(sComp)
                            With Servicing
                                .Log(sComp) = "Serviced Common Component modified: E x p o r t e d !"
                                .Log(sComp) = "Serviced Common Component modified: Re-registered pending release"
                            End With
                            Prgrss.ItemDone = sComp
                        
                        Case Else
                            '~~ Common Component the code has changed, already pending release, possibly conflicting
                            CommonPending.ConflictResolve Serviced.Comp
                            Servicing.Log(sComp) = "Modification conflicting with pending release resolved by update with public common component!"
                            Prgrss.ItemDone = sComp
                    End Select
                End With
nxt:            Prgrss.Dsply
                Set Comp = Nothing
            End If
        Next vbc
    End With
    
    CommonPending.Info
    
    Set Prgrss = Nothing
    mHskpng.RemoveServiceRemains ' remove Temp folder, remove renamed components
    mCompManMenuVBE.Setup
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mExport." & sProc
End Function

Private Function AnExportFolderExists(ByVal oef_path As String, _
                                      ByRef oef_fld As Folder) As Boolean
' --------------------------------------------------------------------------
' When a folder exists with files *.bas, *.cls, or *.frm the function
' returns TRUE and the identified folder (oef_flf).
' --------------------------------------------------------------------------
    Const PROC = "AnExportFolderExists"
    
    On Error GoTo eh
    Dim fld As Folder
    Dim fle As File
    
    For Each fld In fso.GetFolder(oef_path).SubFolders
        For Each fle In fld.Files
            Select Case fso.GetExtensionName(fle.Path)
                Case "bas", "cls", "frm"
                    Set oef_fld = fld
                    AnExportFolderExists = True
                    GoTo xt
            End Select
        Next fle
    Next fld
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub ResolveExportConflict(ByVal c_comp As clsComp)
' ------------------------------------------------------------------------------
' Displays a dialog providing various options regarding a conflicting export
' for a serviced public Common Component which is yet not registered pending
' relase and is not up-to-date.
' ------------------------------------------------------------------------------
    Const PROC = "ResolveExportConflict"
    
    On Error GoTo eh
    Dim Msg                                     As mMsg.udtMsg
    Dim sBttnIgnore                             As String
    Dim sBttnUndo                               As String
    Dim sBttnUpdate                             As String
    Dim sComp                                   As String
    Dim sTitle                                  As String
    
    sBttnIgnore = "Ignore the incident"
    sBttnUndo = "Undo the modifications"
    sBttnUpdate = "Update to current" & vbLf & "public version"
    
    sComp = c_comp.CompName
    sTitle = "Modification of a no up-to-date Common Component!"
    
    With Msg
        With .Section(1)
            .Text.Text = "The export service has detected a conflicting code modification made in the Common Component  " & _
                         mBasic.Spaced(sComp) & "  ." & vbLf & _
                         "This is a serious incident because the modification has been made on a non up-to-date " & _
                         "version of the component. Below is the information about the current public code version of the " & _
                         "component."
        End With
        With .Section(2)
            .Label.Text = "The public version's last modification:"
            .Label.FontColor = rgbBlue
            .Text.MonoSpaced = True
            .Text.Text = "in Workbook : " & c_comp.PublicLastModIn & vbLf & _
                         "on Computer : " & c_comp.PublicLastModOn & vbLf & _
                         "by User     : " & c_comp.PublicLastModBy & vbLf & _
                         "at Date/Time: " & c_comp.PublicLastModAt
        End With
        With .Section(3)
            .Label.Text = "This component's last updated code version:"
            .Label.FontColor = rgbBlue
            .Text.MonoSpaced = True
            .Text.Text = "in Workbook : " & c_comp.ServicedLastModIn & vbLf & _
                         "on Computer : " & c_comp.ServicedLastModOn & vbLf & _
                         "by User     : " & c_comp.ServicedLastModBy & vbLf & _
                         "at Date/Time: " & c_comp.ServicedLastModAt
        End With
        With .Section(4)
            .Label.Text = sBttnIgnore & ":"
            .Label.FontColor = rgbRed
            .Text.Text = "Ignoring will export the modifications but will not register them pending release! " & _
                         "This will not really solve then conflict. This dialog will  popup again when the outdated component is modified." & vbLf & _
                         "Note: Code modifications on a non up-to-date code version or will never become public and thus are a potentially lost effort."
        End With
        With .Section(5)
            .Label.Text = sBttnUndo & ":"
            .Label.FontColor = rgbRed
            .Text.Text = "Discards the recent modifications by re-importing the last Export-File. " & _
                         "This will really solve the conflict since the component will be exported but not registered pending release! " & vbLf & _
                         "Note: Code modifications on a non up-to-date code version or will never become public and thus are a potentially lost effort."
        End With
        With .Section(6)
            .Label.Text = Replace(sBttnUpdate, vbLf, " ") & ":"
            .Label.FontColor = rgbBlue
            .Text.Text = "Updates the component with the current public code version by re-importing " & _
                         "the public Common Component's Export-File from the Common-Components folder. " & _
                         "This is by far the best way to solve the conflict. The component will become ready for being modified " & _
                         "and any subsequent modifications will be registered pending release and may be released to public."
        End With
    End With
    
    Do
        Select Case mMsg.Dsply(d_title:=sTitle _
                             , d_msg:=Msg _
                             , d_label_spec:="R130" _
                             , d_width_min:=40 _
                             , d_buttons:=mMsg.Buttons(mDiff.PublicVersusPendingReleaseBttn(sComp), _
                                                           mDiff.ServicedExportVersusPublicBttn, _
                                                           vbLf, sBttnIgnore, sBttnUndo, sBttnUpdate))
            Case mDiff.PublicVersusPendingReleaseBttn(sComp):   mDiff.PublicVersusPendingReleaseDsply c_comp
            Case mDiff.ServicedExportVersusPublicBttn:          mDiff.ServicedExportVersusPublicDsply c_comp
            Case sBttnUpdate
                '~~ Reset the component to its current public code version
                mUpdate.ByReImport sComp, CommonPublic.LastModExpFile(sComp)
                Servicing.Log(sComp) = "Serviced Common Component's code modifications discarded by re-importing the " & _
                                      "current public version's Export-File from the Common-Components folder"
                Exit Do
            Case sBttnUndo
                '~~ Reset the component to its content before the modification, i.e.
                '~~ the last regular exported code version - which is outdated however.
                mUpdate.ByReImport sComp, c_comp.ExpFileFullName
                Servicing.Log(sComp) = "Serviced Common Component's code modifications reset by re-importing the " & _
                                      "last Export-File (component remains outdated however)"
                Exit Do
            Case sBttnIgnore
                With c_comp
                    .VBComp.Export .ExpFileFullName
                    .SetServicedProperties
                End With
                Servicing.Log(sComp) = "Serviced Common Component exported but not registered pending release!"

                Exit Do
        End Select
    Loop
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

