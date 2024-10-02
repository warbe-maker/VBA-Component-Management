Attribute VB_Name = "mHskpngCommon"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mHskpngCommon:
' ------------------------------------------------------------------------------

Private Sub CompsServicedWithManageGaps()
' ------------------------------------------------------------------------------
' Handles public Common Components which apear having been modified in the
' serviced Wokbook while "out of management" or are serviced for the first time
' but do have outdated Common Components in use. For not confirmed updates the
' export of the corresponding Common Component is suspended in order not to
' obstruct the integrity by handling an outdate Common Component's export as
' regular code modification.
' ------------------------------------------------------------------------------
    Const PROC = "CompsServicedWithManageGaps"
    
    Const BttnConfirmed As String = "Update Confirmed"
    Const BttnDsplyDiff As String = "Display Code Difference"
    Const BttnDisconfirmed  As String = "Continue without confirmation"
   
    On Error GoTo eh
    Dim dct     As Dictionary
    Dim Comp    As clsComp
    Dim Msg     As udtMsg
    Dim v       As Variant
    Dim sLastMod    As String
    
    mBasic.BoP ErrSrc(PROC)
    Set dct = Serviced.CompsCommon
    For Each v In dct
        Set Comp = dct(v)
        With Comp
            If Not mFact.CommCompHasEffectiveManageGap(Comp, sLastMod) Then
                '~~ The used Common Component's code is identical with the current public Common Coponent's code
                If .ExportDisabled(v) Then
                    .ExportDisabled(v) = False
                    .Export
                End If
                GoTo nx
            End If
            
            Msg.Section(1).Text.Text = "The code of the used/hosted Common Component """ & v & """ is not identical " & _
                                       "with the last Export file's content or the Export file doesn't exist. " & _
                                       "I.e. the Workbook is either serviced for the first time with an outdated " & _
                                       "used Common Component or the Workbook has a service gap and the public " & _
                                       "Common Copoent's code has been modified meanwhile."
            
            With Msg.Section(2)
                With .Label
                    .FontColor = rgbBlue
                    .Text = BttnConfirmed
                End With
                .Text.Text = "Strictly recommended. In case not, the export for this Common Component will be suspended " & _
                             "in order not to possible obstruct the public Common Components integrity."
            End With
            With Msg.Section(3)
                With .Label
                    .FontColor = rgbBlue
                    .Text = BttnDisconfirmed
                End With
                .Text.Text = "Not recommended! The Workbook will continue working with an outdated Common Component " & _
                             "and the export for this Common Component will be suspended in order not to possibly " & _
                             "obstruct the public Common Component's integrity."
            End With
            With Msg.Section(4)
                With .Label
                    .FontColor = rgbBlue
                    .Text = BttnDsplyDiff
                End With
                .Text.Text = "Displays the code difference between the currently used Common Component's code " & _
                             "and the current public Common Component's code (Export file in the Common-Components " & _
                             "folder)"
            End With
            
            Do
                Select Case mMsg.Dsply(dsply_title:="Management gap or first time serviced!" _
                                     , dsply_msg:=Msg _
                                     , dsply_Label_spec:="L60" _
                                     , dsply_buttons:=mMsg.Buttons(BttnConfirmed, BttnDsplyDiff))
                    Case BttnConfirmed: .CodeCurrent.ReplaceWith .CodePublic
                    Case BttnDsplyDiff: .CodeCurrent.DsplyDiffs "CodeCurrent", "Current code of the  u s e d  Common Component" _
                                                               , .CodePublic _
                                                               , "CodePublic", "Current code of the  p u b l i c  Common Component"
                    Case Else:          .ExportDisabled(v) = True
                End Select
            Loop
            Prgrss.ItemDone = .CompName
        End With

nx: Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

'Public Sub CompsPublic()
'' ------------------------------------------------------------------------------
'' Removes obsolete sections which are those neither representing an existing
'' VBComponent nor another valid section's Name.
'' ------------------------------------------------------------------------------
'    Const PROC = "CompsPublic"
'
'    On Error GoTo eh
'    Dim dctFiles    As Dictionary
'    Dim dctComps    As Dictionary
'    Dim wbk         As Workbook
'    Dim fld         As Folder
'    Dim fls         As New Dictionary
'    Dim fl          As File
'    Dim sName       As String
'    Dim sExt        As String
'
'    mBasic.BoP ErrSrc(PROC)
'    Set wbk = Serviced.Wrkbk
'
'    If Services Is Nothing Then
'        Set Services = New clsServices
'        Services.Initiate i_service_proc:=mCompManClient.SRVC_RELEASE_PENDING _
'                        , i_wbk_serviced:=ThisWorkbook
'    End If
'
'    Set dctFiles = CommonPublic.Files
'    Set dctComps = CommonPublic.Comps
'    CompsPublicObsolete dctFiles, dctComps
'    CompsPublicMissing dctFiles, dctComps
'
'xt: mBasic.EoP ErrSrc(PROC)
'    Exit Sub
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Public Sub CompsServicedUsedOrPrivate()
' ----------------------------------------------------------------------------
' Ensures that all components with a name known as a public Common Component
' is registered either as used or as private. For any yet not registered
' possible public Common Component a dialog is displayed with the two options.
' ----------------------------------------------------------------------------
    Const PROC = "CompsServicedUsedOrPrivate"
    
    On Error GoTo eh
    Dim BttnPrivate As String
    Dim BttnUsed    As String
    Dim Comp        As clsComp
    Dim dct         As Dictionary
    Dim Msg         As mMsg.udtMsg
    Dim v           As Variant
    Dim wbk         As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    BttnUsed = "Yes!" & vbLf & _
                        "This is a used Common Component" & vbLf & _
                        "identical with the corresponding" & vbLf & _
                        "VBComponent's Export-File in the" & vbLf & _
                        """Common-Components folder"""
    BttnPrivate = "No!" & vbLf & _
                           "This is a VBComponent which just" & vbLf & _
                           "by chance has the same name."
    
    Set dct = Serviced.CompsCommon
    For Each v In dct
        Set Comp = dct(v)
        With Comp
            '~~ The component is a known Common Component
            Select Case CompManDat.RegistrationState(.CompName)
                Case enRegStatePrivate, enRegStateUsed, enRegStateHosted
                    Prgrss.ItemSkipped
                Case Else
                    '~~ Once an equally named VBComponent is registered as private it will no longer be regarded as "used" and therefore not updated
                    '~~ when the corresponding Common Component has been modified.
                    Msg.Section(1).Text.Text = "The component in the VBProject named   " & mBasic.Spaced(.CompName) & "   is known as a ""Common Component"", " & _
                                               "i.e. an equally named component exists in the ""Common-Components folder""  '" & _
                                               CommonPublic.FolderPath & "', but the component is yet neither registered/known as ""used"" nor as ""private"" !" & vbLf & _
                                               "Just a hint by the way: The component may as well be claimed ""hosted"" by this Workbook in case it is yet not " & _
                                               "claimed ""hosted"" by another Workbook/VBProject. *)"
                        
                    With Msg.Section(2)
                        .Label.Text = "*)"
                        With .Text
                             .Text = "See README, section ""Enabling the services (serviced or not serviced)"""
                             .OnClickAction = "https://github.com/warbe-maker/VBA-Component-Management#enabling-the-services-serviced-or-not-serviced"
                             .FontColor = rgbBlue
                        End With
                    End With
                    
                    Select Case mMsg.Dsply(dsply_title:="Not yet registered ""Common Component""" _
                                         , dsply_msg:=Msg _
                                         , dsply_Label_spec:="R25" _
                                         , dsply_buttons:=mMsg.Buttons(BttnUsed, BttnPrivate))
                        Case BttnUsed:     CompManDat.RegistrationState(.CompName) = enRegStateUsed
                                                    CompManDat.LastModAtDateTimeUTC(.CompName) = vbNullString ' yet unknown will force update when outdated
                        Case BttnPrivate:  CompManDat.RegistrationState(.CompName) = enRegStatePrivate
                    End Select
                    Prgrss.ItemDone = .CompName
            End Select
            Prgrss.ItemDone = .CompName
        End With
        Set Comp = Nothing
    Next v

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Components()

    Const PROC = "Components"
    
    On Error GoTo eh
    Dim dctFiles As Dictionary
    Dim dctComps As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Set Prgrss = New clsProgress
    With Prgrss
        .Operation = "Housekeeping Common Components"
        .ItemsTotal = CommonPublic.Comps.Count _
                    + CommonPublic.Files.Count _
                    + CompManDat.Components.Count _
                    + Serviced.Hosted.Count _
                    + (Serviced.CompsCommon.Count * 2) _
                    + CommonPending.Components.Count
        .Figures = False
        .DoneItemsInfo = False
    End With

    Set dctFiles = CommonPublic.Files
    Set dctComps = CommonPublic.Comps
    CompsPublicObsolete dctFiles, dctComps
    CompsPublicMissing dctFiles, dctComps
    
    CompsServicedObsolete       ' CompManDat.Components.Count
    CompsServicedHosted         ' Serviced.Hosted.Count
    CompsServicedUsedOrPrivate  ' Serviced.CompsCommon.Count
    CompsServicedWithManageGaps ' Serviced.CompsCommon.Count
    CompsServicedPendingRelease ' CommonPending.Components.Count

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mHskpngCommon." & sProc
End Function

Private Sub CompsPublicMissing(ByVal h_files As Dictionary, _
                               ByVal h_comps As Dictionary)
' ------------------------------------------------------------------------------
' Maintains the consistency between entries in the CommComps.dat file and the
' exiting components in the Common-Components folder by adding a section for
' an existing Export-File but still missing section, whereby the missing
' information about the origin of the Export-File is taken into account by
' documenting "u n k n o w n !" for the Workbook's name/full-name.
' Note: A missing section indicates a Common Component of wich the Export-File
'       has obvously been copied manually into the Common-Components folder
'       which now in the sense of CompMan has become an available public Common
'       Component ready for being imported into any VB-Project. A new registered
'       Common Component remains "an orphan" until either a Workbbok claims
'       hosting it or simply using and modifying it.
' ------------------------------------------------------------------------------
    Const PROC = "CompsPublicMissing"
    
    On Error GoTo eh
    Dim v         As Variant
    Dim sCompName   As String
    
    mBasic.BoP ErrSrc(PROC)
    For Each v In h_files
        If Not h_comps.Exists(v) Then
            '~~ The Export-File is yet not registered as a known Common Component
            '~~ It may have been copied manually into the Common-Components folder.
            '~~ I.e. its origin is unknown and thus registered as such.
            '~- The origin will remain unknown until the Common Component is
            '~~ modified in a Workbook using or hosting it and exported.
            CommonPublic.LastModInWbk(sCompName) = Nothing
            CommonPublic.LastModAtDateTimeUTC(sCompName) = CompMan.UTC
        End If
        Prgrss.ItemDone = v
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CompsPublicObsolete(ByVal h_files As Dictionary, _
                                ByVal h_comps As Dictionary)
' ------------------------------------------------------------------------------
' Removes in the Common-Component folder's CommComps.dat file any section of
' which no corresponding Export-File exists in the folder
' ------------------------------------------------------------------------------
    Const PROC = "CompsPublicObsolete"
    
    On Error GoTo eh
    Dim sSection    As String
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
        
    '~~ Remove any component of which the corresponding Export-File
    '~~ not exists in the Common-Components folder
    For Each v In h_comps
        If Not h_files.Exists(v) And v <> "@NamesHousekeeping" Then
            CommonPublic.PPFile.SectionRemove sSection
        End If
        Prgrss.ItemDone = v
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CompsServicedHosted()
' ------------------------------------------------------------------------------
' Maintains in the serviced Workbook's CompMan.dat file for any component
' claimed hosted the corresponding RegistrationState and the Revision-Number
' (when missing).
' ------------------------------------------------------------------------------
    Const PROC      As String = "CompsServicedHosted"
    
    On Error GoTo eh
    Dim wbk         As Workbook
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Serviced.Wrkbk
    
    With CompManDat
        For Each v In Serviced.Hosted
            If Serviced.CompsCommon.Exists(v) Then
                .RegistrationState(v) = enRegStateHosted
            Else
                Err.Raise AppErr(1), ErrSrc(PROC), _
                "The serviced Workbook """ & Serviced.Wrkbk.Name & """ claims the component """ & v & _
                """ hosted but a component with this name does not exist in the Workbook's VBProject!"
            End If
            Prgrss.ItemDone = v
        Next v
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CompsServicedNotHosted()
' ----------------------------------------------------------------------------
' When a former hosting Workbook not or no longer claims a Common Component
' hosted the RegistrationState in the serviced Workbook's CompMan.dat file
' is changed to enRegStateUsed.
' ----------------------------------------------------------------------------
    Const PROC      As String = "CompsServicedNotHosted"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.ServicedWbk
    For Each v In CompManDat.Components
        If Not Serviced.Hosted.Exists(v) Then
            If mComp.Exists(v, wbk) Then
                CompManDat.RegistrationState(v) = enRegStateUsed
            End If
        End If
        Prgrss.ItemDone = v
    Next v

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CompsServicedObsolete()
' ------------------------------------------------------------------------------
' Remove sections representing VBComponents no longer existing and those with an
' invalid name.
' ------------------------------------------------------------------------------
    Const PROC = "CompsServicedObsolete"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    With CompManDat
        For Each v In .Components
            If Not .IsSystemSection(v) Then
                If Not Serviced.CompsCommon.Exists(v) Then
                    '~~ The section no longer represents an existing component
                    .ComponentRemove v
                End If
            End If
            Prgrss.ItemDone = v
        Next v
    End With
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

'Public Sub CompsServiced()
'' ------------------------------------------------------------------------------
'' Performs a houskeeping for the serviced workbook's CompMan.dat Private Profile
'' file by
'' - removing obsolete entries (sections not representing a VB-Project component
'' - maintaining sections representing VB-Project components known as Common
''   Components (those existing in the CommComps.dat file or being hosted.
'' ------------------------------------------------------------------------------
'    Const PROC = "CompsServiced"
'
'    On Error GoTo eh
'    mBasic.BoP ErrSrc(PROC)
'
'    CompsServicedObsolete       ' CompManDat.Components.Count
'    CompsServicedHosted         ' Serviced.Hosted.Count
'    CompsServicedUsedOrPrivate  ' Serviced.CompsCommon.Count
'    CompsServicedWithManageGaps ' Serviced.CompsCommon.Count
'    CompsServicedPendingRelease ' CommonPending.Components.Count
'
'xt: mBasic.EoP ErrSrc(PROC)
'    Exit Sub
'
'eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Private Sub CompsServicedPendingRelease()
' ------------------------------------------------------------------------------
' Checks a serviced Workbooks Common Component's last modification against the
' current public Common Component and against the registered pending release
' components. Displays an appropriate dialog in case of any discrepancies.
' ------------------------------------------------------------------------------
    Const PROC = "CompsServicedPendingRelease"
    
    On Error GoTo eh
    Dim Comp                        As clsComp
    Dim sPublicLastModDateTimeUtc   As String
    Dim sPublicLastModInWbk         As String
    Dim sPublicLastModOnMachine     As String
    Dim sPendingLastModDateTimeUtc  As String
    Dim sPendingLastModInWbk        As String
    Dim sPendingLastModOnMachine    As String
    Dim sServicedLastModDateTimeUtc As String
    Dim sServicedLastModInWbk       As String
    Dim sServicedLastModOnMachine   As String
    Dim v                           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    For Each v In CommonPending.Components
        sServicedLastModDateTimeUtc = CompManDat.LastModAtDateTimeUTC(v)
        If CommCompIsPendingRelease(v, sPendingLastModDateTimeUtc _
                                     , sPendingLastModInWbk _
                                     , sPendingLastModOnMachine) Then
            If sPendingLastModDateTimeUtc <> sServicedLastModDateTimeUtc Then
                Debug.Print ErrSrc(PROC) & ": " & v & ": The last modification of """ & v & """ is not the pending one!"
                Debug.Print ErrSrc(PROC) & ": " & String(Len(v & ": "), " ") & Serviced.Wrkbk.Name & ": " & sServicedLastModDateTimeUtc
                Debug.Print ErrSrc(PROC) & ": " & String(Len(v & ": "), " ") & FSo.GetFileName(sPendingLastModInWbk) & ": " & sPendingLastModDateTimeUtc
                Set Comp = Serviced.CompsCommon(v)
            End If

'            If mFact.IsCommCompPublic(Comp, sPublicLastModDateTimeUtc _
'                                          , sPublicLastModInWbk _
'                                          , sPublicLastModOnMachine) Then
'            End If
        Else
            If sServicedLastModDateTimeUtc > CommonPublic.LastModAtDateTimeUTC(v) Then
                '~~ A modification has yet not been registered pending release
                Set Comp = New clsComp
                With Comp
                    .Wrkbk = Serviced.Wrkbk
                    .CompName = v
                End With
                If CommonPending Is Nothing Then Set CommonPending = New clsCommonPending
                CommonPending.Manage Comp
                Set Comp = Nothing
            End If
        End If
        Prgrss.ItemDone = v
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


