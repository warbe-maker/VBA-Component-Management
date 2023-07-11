Attribute VB_Name = "mHskpng"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mHskpng." & s
End Function

Public Sub CommComps(ByVal h_hosted As String)
' ------------------------------------------------------------------------------
' Removes obsolete sections which are those neither representing an existing
' VBComponent no another valid section's Name.
' ------------------------------------------------------------------------------
    Const PROC = "CommComps"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mHskpng.CommCompsMaintainProperties
    mHskpng.CommCompsRemoveObsoleteComponents h_hosted
    mHskpng.CommCompsHosted h_hosted
    mHskpng.CommCompsNotHosted h_hosted
    mHskpng.CommCompsUsed
    mHskpng.ReorgDatFile CommCompsDatFileFullName
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ReorgDatFile(ByVal r_dat_file As String)
    Const PROC = "ReorgDatFile"
    mBasic.BoP ErrSrc(PROC)
    mFso.PPreorg r_dat_file
    mBasic.EoP ErrSrc(PROC)
End Sub

Private Sub CommCompsMaintainProperties()
' ------------------------------------------------------------------------------
' Adds for each Common Component's Export-File in the Common-Components folder
' a section to the CommComps.dat when missing or updates the ExportFileName when
' not identical with the found file.
'
' Background:
' A missing section indicates a Common Component of wich the Export-File has
' obvously been copied manually into the Common-Components folder which now in
' the sense of CompMan has become an available Common Component ready for being
' imported into any VB-Project. A new registered Common Component remains
' un-hosted until a Workbbok claims hosting it, i.e. providing a delevelopment
' and test environment for it.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsMaintainProperties"
    
    On Error GoTo eh
    Dim fle         As File
    Dim fso         As New FileSystemObject
    Dim dct         As Dictionary
    Dim sCompName   As String
    Dim sExt        As String
    
    mBasic.BoP ErrSrc(PROC)
    Set dct = mCommComps.Components
    With fso
        For Each fle In .GetFolder(wsConfig.FolderCommonComponentsPath).Files
            sExt = .GetExtensionName(fle.Path)
            Select Case sExt
                Case "bas", "frm", "cls"
                    sCompName = .GetBaseName(fle.Path)
                    If Not dct.Exists(sCompName) Then
                        '~~ The Export-File is yet not registered as a known Common Component
                        '~~ It most likely has been copied manually into the Common-Components
                        '~~ folder. I.e. its "raw host" is unknown - and registered as such.
                        '~- The raw host will remain unknown until the Common Component is
                        '~~ modified in a Workbook using it and exported.
                        mCommComps.LastModWbk(sCompName) = Nothing
                        mCommComps.RevisionNumber(sCompName) = CompManDat.RevisionNumberInitial
                    Else
                        If mCommComps.LastModExpFileFullNameOrigin(sCompName) = vbNullString Then
                            Debug.Print "The property ""LastModExpFileFullNameOrigin"" of component " & sCompName & " is not available, i.e. its origin is unknown or simply yet not registered respectively!"
                        End If
                    End If
            End Select
        Next fle
    End With
    
    Set fso = Nothing
    Set dct = Nothing

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsRemoveObsoleteComponents(ByVal h_hosted As String)
' ------------------------------------------------------------------------------
' Remove in the PrivateProfile file CommComps.dat:
' - Sections representing VBComponents for which an Export-File does not exist
'   in the Common-Components folder
' - Sections indicating a Common Component of the serviced Workbook but the
'   component not exists.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsRemoveObsoleteComponents"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim wbk         As Workbook
    Dim dct         As Dictionary
    Dim fso         As New FileSystemObject
    Dim sBaseName   As String
    Dim dctHosted   As Dictionary
    Dim sExpFile    As String
    
    mBasic.BoP ErrSrc(PROC)
    Set dctHosted = Hosted(h_hosted)
    Set wbk = Services.Serviced
    sBaseName = fso.GetBaseName(wbk.FullName)
    Set dct = mCommComps.Components
    
    '~~ Obsolete because the component is no longer hosted by the indicated Workbook
    '~~ no longer exist in the indicated Workbook
    For Each v In dct
        If mCommComps.LastModWbkBaseName(v) = sBaseName Then
            '~~ The component indicates being one of the serviced Workbook
            If Not mComp.Exists(v, wbk) Then
                CompManDat.RemoveComponent v
            End If
        End If
    Next v
    
    '~~ Obsolete because the corresponding Export-File
    '~~ no longer exists in the Common-Components folder
    '~~ De-register global Common Components no longer hosted
    Set dct = mCommComps.Components
    For Each v In dct
        sExpFile = fso.GetFileName(mCommComps.LastModExpFileFullNameOrigin(v))
        If sExpFile <> vbNullString Then
            If Not fso.FileExists(wsConfig.FolderCommonComponentsPath & "\" & sExpFile) Then
                CommCompsRemoveSection v
            End If
        End If
    Next v
    Set dct = mCommComps.Components
    
xt: Set dct = Nothing
    Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsUsed()
' ----------------------------------------------------------------------------
' Manages the registration of used Common Components, done before change
' components are exported and used changed Common Components are updated.
' When not yet registered a confirmation dialog ensures a component not just
' accidentially has the same name. The type of confirmation is registered
' either as "used" or "private" together with the current revision number.
' When none is available the current date is registered on the fly.
' its Revision-Number is
' ----------------------------------------------------------------------------
    Const PROC          As String = "CommCompsUsed"
    Dim vbc             As VBComponent
    Dim wbk             As Workbook
    Dim BttnConfirmed   As String
    Dim BttnPrivate     As String
    Dim Msg             As mMsg.TypeMsg
    Dim Comp            As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    BttnConfirmed = "Yes!" & vbLf & _
                    "This is a used Common Component" & vbLf & _
                    "identical with the corresponding" & vbLf & _
                    "VBComponent's Export-File in the" & vbLf & _
                    """Common-Components folder"""
    BttnPrivate = "No!" & vbLf & _
                  "This is a VBComponent which" & vbLf & _
                  "accidentially has the same name."
    
    Set wbk = Services.Serviced
    For Each vbc In wbk.VBProject.VBComponents
        Set Comp = New clsComp
        With Comp
            .Wrkbk = wbk
            .VBComp = vbc
            If mCommComps.Exists(vbc, .ExpFileExt) Then
                If mCommComps.RevisionNumber(vbc.Name) = vbNullString Then
                    mCommComps.RevisionNumber(vbc.Name) = vbNullString
                End If
                If Not CompManDat.RegistrationState(vbc.Name) = enRegStatePrivate _
                And Not CompManDat.RegistrationState(vbc.Name) = enRegStateUsed _
                And Not CompManDat.RegistrationState(vbc.Name) = enRegStateHosted _
                Then
                    '~~ Once an equally named VBComponent is registered a private it will no longer be regarded as "used" and updated.
                    Msg.Section(1).Text.Text = "The VBComponent named   " & mBasic.Spaced(vbc.Name) & "   is known as a ""Common Component"" " & _
                                               "because it exists in the ""Common-Components folder""  '" & _
                                               wsConfig.FolderCommonComponentsPath & "'  but is yet not registered either " & _
                                               "as used or private in the serviced Workbook."
                    
                    Select Case mMsg.Dsply(dsply_title:="Not yet registered ""Common Component""" _
                                         , dsply_msg:=Msg _
                                         , dsply_buttons:=mMsg.Buttons(BttnConfirmed, vbLf, BttnPrivate))
                        Case BttnConfirmed: CompManDat.RegistrationState(vbc.Name) = enRegStateUsed
                                            CompManDat.RevisionNumber(vbc.Name) = vbNullString ' yet unknown will force update when outdated
                        Case BttnPrivate:   CompManDat.RegistrationState(vbc.Name) = enRegStatePrivate
                    End Select
                End If
            Else
                '~~ The Export-File has manually been copied into the Common
                '~~ Components-Folder and thus is yet not registered
                
            End If
        End With
        Set Comp = Nothing
    Next vbc

xt: mBasic.EoP ErrSrc(PROC)

End Sub

Private Sub CommCompsRemoveSection(ByVal s As String)
    mFso.PPremoveSections CommCompsDatFileFullName, s
End Sub

Private Sub CommCompsHosted(ByVal m_hosted As String)
' ----------------------------------------------------------------------------
' - Registers the Workbook as 'Raw-Host' when it hosts at least one Common
'   Component
' - Maintains an up-to-date copy of the Export-File in the Common-Components
'   folder
' - Maintains for each hosted (raw) Common Component the properties:
'   - in the local CommComps.dat:
'     - Component Name
'     - Revision Number
'   - in the ComComps-RawsSaved.dat in the Common-Components folder:
'     - Component Name
'     - Export File Full Name
'     - Host Base Name
'     - Host Full Name
'     - Host Name
'     - Revision Number
' ----------------------------------------------------------------------------
    Const PROC = "CommCompsHosted"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim fso             As New FileSystemObject
    Dim CommCompHosted  As clsComp
    Dim sHostBaseName   As String
    Dim dctHosted       As Dictionary
    Dim wbk             As Workbook
    Dim sInconstMsg     As String
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    sHostBaseName = fso.GetBaseName(wbk.FullName)
    Set dctHosted = mCommComps.Hosted(m_hosted)
                    
    For Each v In dctHosted
        If Not mComp.Exists(v, wbk) Then
            MsgBox "The VBComponent " & v & " is claimed hosted by the serviced Workbook " & Services.Serviced.Name & _
                   " will be ignored (it does not exist in the Workbook's VB-Project) !" & vbLf & vbLf & _
                   "When the component is no longer hosted or its name has changed the argument needs to be updated accordingly.", _
                   vbOK, "VBComponent " & v & "does not exist!"
        Else
            Set CommCompHosted = New clsComp
            With CommCompHosted
                .Wrkbk = wbk
                .CompName = v
                If CompManDat.RegistrationState(v) <> enRegStateHosted Then
                    '~~ The Workbook has yet not claimed the Common Component hosted but now does.
                    CompManDat.RegistrationState(v) = enRegStateHosted
                    '~~ This housekeeping is executed prior the "Export of changed components"
                    '~~ The Export-File comparison is therefore done with a temporary Export-File.
                    If Not Services.FilesDiffer(f_file_1:=.ExpFileTemp _
                                              , f_file_2:=mCommComps.LastModExpFile(v) _
                                              , f_ignore_export_header:=True) _
                    Then
                        '~~ Only when the claiming Workbook's Common Component is identical
                        '~~ with the Export-File in the Common-Components folder it is also
                        '~~ registered as the raw source
                        mCommComps.LastModWbk(v) = wbk
                        mCommComps.LastModExpFileFullNameOrigin(v) = .ExpFileFullName
                        mCommComps.RevisionNumber = .RevisionNumber
                    Else
                        '~~ Any other Workbook appears to have modified and saved the Common Component
                        '~~ to the Common-Components folde. This information remains valid.
                    End If
                End If
                
                If Not mCommComps.Exists(.VBComp, .ExpFileExt) Then
                    mCommComps.SaveToCommonComponentsFolder v, .ExpFile, .ExpFileFullName, wbk
                    
                End If
                
                If Not Services.FilesDiffer(f_file_1:=.ExpFile _
                                          , f_file_2:=mCommComps.LastModExpFile(v)) Then
                    If mCommComps.RevisionNumber(v) <> CompManDat.RevisionNumber(v) Then
                        mCommComps.RevisionNumber(v) = CompManDat.RevisionNumber(v)
                    End If

                End If
            End With
            Set CommCompHosted = Nothing
        End If
    Next v

xt: Set fso = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsNotHosted(ByVal h_hosted As String)
' ----------------------------------------------------------------------------
' When a former hosting Workbook not or no longer claims a Common Component
' hosted the RegistrationState is changed to enRegStateUsed.
' ----------------------------------------------------------------------------
    Const PROC      As String = "CommCompsNotHosted"
    
    Dim dctHosted   As Dictionary
    Dim dctComps    As Dictionary
    Dim v           As Variant
    Dim wbk         As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    Set dctHosted = mCommComps.Hosted(h_hosted)
    Set dctComps = mCommComps.Components
    For Each v In dctComps
        If LastModWbkName(v) = wbk.Name Then
            If Not dctHosted.Exists(v) Then
                If mComp.Exists(v, wbk) Then
                    CompManDat.RegistrationState(v) = enRegStateUsed
                End If
            End If
        End If
    Next v

xt: mBasic.EoP ErrSrc(PROC)
    
End Sub



