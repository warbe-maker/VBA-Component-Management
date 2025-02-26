Attribute VB_Name = "mHskpng"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mHskpng: Provides a consistent, up-to-date, and valid data
' ======================== base regarding Common Components, specifically by
' considering possible manual interventions by the developer such as manually
' copying an Export-File to the Common-Components folder to make it public for
' instance.
' Addresses:
' - History environment file/folder names forwarded to the current names
'   specified
' - Possible manual user interactions such like copying a Common Component
'   Export-Files to the Common-Components folder, thereby bypassing a registration
'   pending release and subsequently using the release to public service.
' - Possible new imported public Common Components from the Common-Components
'   folder
' - Possible copies of Common Components from one VB-Project into another.
'
' Public Properties:
' ------------------
' CompManServiceFolder
' ExportFolder              Up-to-date Export-Folder name, also in case the
'                           configured name had changed), including an
'                           up-to-date content, i.e. any Export-File not
'                           corresponding with an existing component removed.
' ServicesLogFileFullName
' ServicesTraceLogFileFullName
' CommCompsServicedPrivProfFileFullName
'
' Public services:
' ----------------
' FocusOnOpen               Maintains an up-to-date data base regarding Common
'                           Components when the serviced Workbook is opened.
' FocusOnSave               Maintains an up-to-date data base regarding Common
'                           Components when the serviced Workbook is saved.
' RemoveServiceRemains      Removes any temporary export files created by the
'                           update service and any components renamed by the
'                           update service.
'
' ------------------------------------------------------------------------------

Private Sub CommCompsPublicNew(ByVal c_files As Dictionary, _
                               ByVal c_comps As Dictionary, _
                      Optional ByVal c_sequ_no As Long = 0)
' ------------------------------------------------------------------------------
' Maintains the consistency between existing Export-Files in the Common-
' Components folder and entries/sections in the CommComps.dat file. Any missing
' entry indicates a manually added Export-File. In case a section for the new
' existing Export-File is added. When the Export-File meets the corresponding
' Export-File of the serviced Workbook, thisone is regarded the origin Export-
' File and registered as such.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsPublicNew"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim sExpFile    As String
    Dim sCompName   As String
    
    mBasic.BoP ErrSrc(PROC)
    With CommonPublic
        For Each v In c_files
            sCompName = v
            sExpFile = c_files(v)
            If Not c_comps.Exists(sCompName) Then
                '~~ The Export-File is yet not registered as a known Common Component
                '~~ It may have been copied manually into the Common-Components folder.
                '~~ I.e. its origin is either the Serviced Workbooks corresponding component
                '~~ or it is unknown. In the latter case the Export-File in the Common-Components
                '~~ folder itself is reand thus registered as such.
                '~- The origin will remain unknown until the Common Component is
                '~~ modified in a Workbook using or hosting it and exported.
                With Serviced.Comp
                    .CompName = sCompName
                    If .CodeExprtd.Meets(.CodePublic) Then
                        .SetServicedProperties
                        CommonServiced.KindOfComponent(sCompName) = enCompCommonUsed
                        CommonPublic.SetPropertiesEqualServiced (sCompName)
                    Else
                        CommonPublic.LastModExpFileOrigin(sCompName) = sExpFile
                    End If
                End With
            Else
                If fso.GetFileName(CommonPublic.LastModExpFileOrigin(sCompName)) = vbNullString Then
                    CommonPublic.LastModExpFileOrigin(sCompName) = sExpFile
                End If
            End If
            Prgrss.ItemDone(c_sequ_no) = v
        Next v
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsPublicObsolete(ByVal c_files As Dictionary, _
                                    ByVal c_comps As Dictionary, _
                           Optional ByVal c_sequ_no As Long = 0)
' ------------------------------------------------------------------------------
' - Removes in the Common-Component folder's CommComps.dat file any section of
'   which no corresponding Export-File exists in the folder.
' - Removes the corresponding section in the seviced Workbook's CommComps.dat file
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsPublicObsolete"
    
    On Error GoTo eh
    Dim v As Variant
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Remove any component of which the corresponding Export-File
    '~~ not exists in the Common-Components folder
    For Each v In c_comps
        If Not c_files.Exists(v) Then
            With CommonPublic
                .Remove v
                If .Components.Exists(v) Then .Components.Remove v
                CommonServiced.RemoveComponent v
                If CommonServiced.Components.Exists("clsCode") Then Stop
            End With
        End If
        Prgrss.ItemDone(c_sequ_no) = v
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsRemoveRenamedByUpdate()
' ------------------------------------------------------------------------------
' Removes all Common Components which had been renamed in order to enable the
' import of an Export-File (from the Common-Components folder) wich represents
' an up-to-date code version of a Common Component.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsRemoveRenamedByUpdate"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim sComp   As String
    
    With Serviced.Wrkbk.VBProject
        For Each vbc In .VBComponents
            sComp = vbc.Name
            If Serviced.HasBeenRenamedByUpdateService(sComp) Then
                Servicing.Log(sComp) = "Component """ & sComp & """ (resulting from rename to enable import of up-to-date Common Component) removed."
                .VBComponents.Remove vbc
            End If
        Next vbc
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsServicedHosted(Optional ByVal c_sequ_no As Long = 0)
' ------------------------------------------------------------------------------
' Maintains in the serviced Workbook's CommComps.dat file for any component
' claimed hosted the corresponding KindOfComponent and the Revision-Number
' (when missing).
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsServicedHosted"
    
    On Error GoTo eh
    Dim wbk         As Workbook
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Serviced.Wrkbk
    
    With CommonServiced
        For Each v In Serviced.Hosted
            If Serviced.CompExists(v) Then
                .KindOfComponent(v) = enCompCommonHosted
            Else
                Err.Raise AppErr(1), ErrSrc(PROC), _
                "The serviced Workbook """ & Serviced.Wrkbk.Name & """ claims the component """ & v & _
                """ hosted but a component with this name does not exist in the Workbook's VBProject!"
            End If
            Prgrss.ItemDone(c_sequ_no) = v
        Next v
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub CommCompsServicedKindOf(Optional ByVal c_sequ_no As Long = 0)
' ----------------------------------------------------------------------------
' Ensures that all components with a name known as a public Common Component
' are registered either as used or as private. For any yet not registered
' possible public Common Component a dialog is displayed with the two options.
' Note: This procedure is Public testing only.
' ----------------------------------------------------------------------------
    Const PROC = "CommCompsServicedKindOf"
    
    On Error GoTo eh
    Dim bDone       As Boolean
    Dim BttnPrivate As String
    Dim BttnUsed    As String
    Dim sComp       As String
    Dim dct         As Dictionary
    Dim Msg         As mMsg.udtMsg
    Dim v           As Variant
    Dim sKnownAs    As String
    
    mBasic.BoP ErrSrc(PROC)
    BttnPrivate = "No!" & vbLf & _
                  "This is a  p r i v a t e  component" & vbLf & _
                  "(just by chance with the same name)"
    
    Set dct = CommonPublic.All
    For Each v In Serviced.Wrkbk.VBProject.VBComponents
        sComp = v.Name
        If dct.Exists(sComp) Then
            Select Case CommonServiced.KindOfComponent(sComp)
                Case enCompCommonPrivate, enCompCommonUsed, enCompCommonHosted
                Case Else
                    sKnownAs = vbNullString
                    Select Case True
                        Case CommonPublic.Exists(sComp)
                            sKnownAs = "Common Component public"
                            BttnUsed = "Yes!" & vbLf & _
                                       "This is a  u s e d  Common Component" & vbLf & _
                                       "identical with the corresponding" & vbLf & _
                                       "VBComponent's Export-File in the" & vbLf & _
                                       """Common-Components folder"""
                        Case CommonPending.Exists(sComp)
                            sKnownAs = "Common Component ""pending release"" (i.e. it will become public once released)"
                            BttnUsed = "Yes!" & vbLf & _
                                       "This is a  u s e d  Common Component" & vbLf & _
                                       "identical with the corresponding" & vbLf & _
                                       "VBComponent's Export-File in the" & vbLf & _
                                       """PendingReleases"" folder" & vbLf & _
                                       "(a public Common Component once released)"
                    End Select
                    If sKnownAs <> vbNullString Then
                        '~~ The component is known as a public Common Component hosted or pending release
                        With Serviced.Comp
                            .CompName = sComp
                            Select Case CommonServiced.KindOfComponent(sComp)
                                Case enCompCommonPrivate, enCompCommonUsed, enCompCommonHosted
                                Case Else
                                    '~~ Once an equally named VBComponent is registered as private it will no longer be regarded as "used" and therefore not updated
                                    '~~ when the corresponding Common Component has been modified.
                                    Msg.Section(1).Text.Text = "The component   " & mBasic.Spaced(sComp) & "   is known as a """ & sKnownAs & """ " & _
                                                               "but the component is yet neither registered/known as ""used"" nor as ""private"" !" & vbLf & _
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
                                    With Msg.Section(3)
                                        .Label.Text = "About:"
                                        .Label.FontColor = rgbBlue
                                        .Text.Text = mCompMan.AboutCommComps
                                    End With
                                    
                                    Select Case mMsg.Dsply(d_title:="Not yet registered ""Common Component""" _
                                                         , d_msg:=Msg _
                                                         , d_label_spec:="L30" _
                                                         , d_buttons:=mMsg.Buttons(BttnUsed, BttnPrivate))
                                        Case BttnUsed:  CommonServiced.KindOfComponent(.CompName) = enCompCommonUsed
                                                        CommonServiced.LastModAt(.CompName) = vbNullString ' yet unknown will force update when outdated
                                                        If .IsCommCompUpToDate Then
                                                            .Export
                                                            CommonServiced.SetPropertiesEqualPublic .CompName
                                                        End If
                                        Case BttnPrivate:  CommonServiced.KindOfComponent(.CompName) = enCompCommonPrivate
                                    End Select
                            End Select
                        End With
                    End If
            End Select
        End If
        Prgrss.ItemDone(c_sequ_no) = sComp
    Next v

xt:
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsServicedNotHosted(Optional ByVal c_sequ_no As Long = 0)
' ----------------------------------------------------------------------------
' When a former hosting Workbook not or no longer claims a Common Component
' hosted the KindOfComponent in the serviced Workbook's CommComps.dat file
' is changed to enCompCommonUsed.
' ----------------------------------------------------------------------------
    Const PROC      As String = "CommCompsServicedNotHosted"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Servicing.ServicedWbk
    With CommonServiced
        For Each v In .Components
            If Not Serviced.Hosted.Exists(v) Then
                If Serviced.CompExists(v) Then
                    .KindOfComponent(v) = enCompCommonUsed
                End If
            End If
            Prgrss.ItemDone(c_sequ_no) = v
        Next v
    End With

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsServicedObsolete(Optional ByVal c_sequ_no As Long = 0)
'' ------------------------------------------------------------------------------
' Remove sections representing VBComponents no longer existing and those with an
' invalid name.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsServicedObsolete"
    
    On Error GoTo eh
    Dim v As Variant
    
    mBasic.BoP ErrSrc(PROC)
    With CommonServiced
        For Each v In .Components
            If Not .IsSystemSection(v) Then
                If Not CommonPublic.Exists(v) Or Not Serviced.CompExists(v) Then
                    '~~ The section no longer represents an existing Common Component
                    '~~ or the component no longer exists in the VB-Project
                    .ComponentRemove v
                End If
            End If
            Prgrss.ItemDone(c_sequ_no) = v
        Next v
    End With
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsServicedPendingRelease(Optional ByVal c_sequ_no As Long = 0)
' ------------------------------------------------------------------------------
' Maintains a consistent pending releases data base by:
' 1. Removing pending Common Components of which the Export-file is already
'    identical with the current public Common Component's Export-Fie in the
'    Common-Components folder.
' 2. Removing entries in the Pending.dat file without a corresponding Export-
'    file in the Pending folder
' 3. Removing Export-files without a corresponding entry in the Pending.dat file
'    File in the Common-Components\Pending folder
' 4. When the housekeeping runs prior the Export Service and an up-to-date
'    used Common Component is pending release due to a modification in
'    another Workbook a dialog is displayed which proposes the continuation -
'    and possibly finalization in the opened Workbook. This will prevent
'    a possible concurrent modification which will have to be prevented.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsServicedPendingRelease"
    
    On Error GoTo eh
    Dim Comp                        As clsComp
    Dim dctPendingExpFiles          As Dictionary
    Dim dctPublicComps              As Dictionary
    Dim v                           As Variant
    Dim sComp                       As String
    Dim vbc                         As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    '~~ 1. Removing pending Common Components of which the Export-file is identical
    '~~    with the current public Common Component's Export-Fie in the Common-
    '~~    Components folder.
    If CommonPending.Components.Count <> 0 Then
        Set dctPublicComps = CommonPublic.Components
        With Serviced
            For Each v In CommonPending.Components
                sComp = v
                If .CompExists(sComp) Then
                    With .Comp
                        .CompName = sComp
                        If Not .CodePublic.IsNone And Not .CodePnding.IsNone Then
                            '~~ There's already a public Common Component in the Common-Components folder
                            If .CodePublic.Meets(.CodePnding) Then
                                '~~ When the public code is identical with the pending release code the pending is obsolete.
                                '~~ A likely reason is that the export file had been copied manually. The properties
                                '~~ are updated anyway.
                                CommonPublic.SetPropertiesEqualPending sComp
                                CommonPending.Remove sComp
                            End If
                        End If
                    End With
                End If
                Prgrss.ItemDone(c_sequ_no) = v
            Next v
        End With
    End If
                                            
    With CommonPending
        '~~ 2. Remove Export-files without a corresponding entry in the Pending.dat file
        Set dctPendingExpFiles = .ExportFiles
        For Each v In dctPendingExpFiles
            If Not .Exists(v) Then
                fso.DeleteFile dctPendingExpFiles(v)
            End If
            Prgrss.ItemDone(c_sequ_no) = v
        Next v
        If dctPendingExpFiles.Count = 0 And fso.FolderExists(mEnvironment.CommCompsPendingPath) _
        Then fso.DeleteFolder mEnvironment.CommCompsPendingPath
    
        '~~ 3. Remove entries in the Pending.dat file without a corresponding Export-
        '~~    File in the Common-Components\Pending folder - may have been moved manually to
        '~~    the Common-Components folder.
        For Each v In .Components
            sComp = v
            If Not dctPendingExpFiles.Exists(sComp) Then
                With New clsComp
                    .CompName = v
                    CommonPublic.SetPropertiesEqualPending sComp
                End With
                CommonPending.Remove sComp
            End If
            Prgrss.ItemDone(c_sequ_no) = v
        Next v
    End With
    
    '~~ 4. Regard pending release Common Components of which the the code is
    '~~    identical with the corresponding serviced Workbook's component as
    '~~    "pending release by the serviced Workbook".
    For Each v In CommonPending.Components
        sComp = v
        If Serviced.CompsCommon.Exists(sComp) Then
            Set Comp = New clsComp
            With Comp
                .CompName = sComp
                Select Case True
                    Case .CodePnding.Meets(.CodeCrrent)
                        .Export
                        CommonPending.SetPropertiesEqualServiced sComp
                    Case fso.FileExists(.ExpFileFullName)
                        If .CodePnding.Meets(.CodeExprtd) Then
                            CommonPending.SetPropertiesEqualServiced sComp
                        End If
                End Select
            End With
        End If
        Prgrss.ItemDone(c_sequ_no) = v
    Next v
    
    '~~ Propose continuation of modification in opened serviced Workbook
    For Each v In CommonPending.Components
        sComp = v
        With CommonServiced
            If .KindOfComponent(sComp) = enCompCommonUsed Or .KindOfComponent(sComp) = enCompCommonHosted Then
                Select Case True
                    Case CommonPending.LastModInWrkbkFullName(sComp) = Serviced.Wrkbk.FullName _
                      And CommonPending.LastModOn(sComp) = mEnvironment.ThisComputersName
                    Case Else
                        '~~ Propose switch
                        ProposeContinuationOfModificationInThisWorkbook sComp
                End Select
            End If
        End With
        Prgrss.ItemDone(c_sequ_no) = v
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsServicedProperties(Optional ByVal c_sequ_no As Long = 0)
' ------------------------------------------------------------------------------
' Maintains for all serviced Common Components the properties in the serviced
' workbook's CommComps.dat file - specifically when different from the public
' versions properties differ from the serviced component's properties although
' the code is identical with the public version.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsServicedProperties"
    
    On Error GoTo eh
    Dim bDone   As Boolean
    Dim sComp   As String
    Dim v       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Loop through all components in the serviced Workbook considered a Commom Component
    With Serviced
        For Each v In CommonPublic.Components
            sComp = v
            If .CompExists(sComp) Then
                .Comp.CompName = sComp
                With .Comp
                    If .CodeCrrent.Meets(.CodePublic) Then
                        '~~ A uses Common Component's code is identical with the public version
                        If Not .CodeCrrent.Meets(.CodeExprtd) Then
                            '~~ Apparently the public Common Component's Export-File has been imported by the VBE
                            '~~ and either has yet not exported or the export file is outdated
                            .Export
                            With Servicing
                                .Log(sComp) = "Serviced Common Component properties housekeeping:"
                                .Log(sComp) = "Serviced Common Component (public assumed imported manually): e x p o r t e d !"
                                .Log(sComp) = "Serviced Common Component (public assumed imported manually): Properties set equal public"
                            End With
                        End If
                        CommonServiced.SetPropertiesEqualPublic sComp, bDone
                        If bDone Then
                            With Servicing
                                .Log(sComp) = "Serviced Common Component properties housekeeping:"
                                .Log(sComp) = "Serviced Common Component's Export-File meets public: Properties set equal public"
                            End With
                        End If
                    End If
                End With
            End If
nx:         Prgrss.ItemDone(c_sequ_no) = v
        Next v
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ServiceOutstandingAtOpen(Optional ByVal s_sequ_no As Long = 0) As Long
' ------------------------------------------------------------------------------
' Performs an Export when yet not/never done or returns kind of outstanding
' service as a numeric value:
' 1 A modification had been made on a Common Component while out-of-service but
'   it conflicts with one meanwhile made in another VB-Project.
' ------------------------------------------------------------------------------
    Const PROC = "ServiceOutstandingAtOpen"
    
    On Error GoTo eh
    Dim lKind   As enKindOfComp
    Dim vbc     As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Loop through all components in the serviced Workbook considered a Commom Component
    With Serviced
        For Each vbc In .Wrkbk.VBProject.VBComponents
            With .Comp
                .CompName = vbc.Name
                lKind = CommonServiced.KindOfComponent(.CompName)
                Select Case True
                    Case Not fso.FileExists(.ExpFileFullName) _
                        '~~ Done with export, may be outdated though if a Common Component
                        .Export
                        ServiceOutstandingAtOpen = 0
                    
                    '~~ Export-File exists cases
                    Case Not .CodeExprtd.DiffersFrom(.CodeCrrent) ' no service outstanding
                        ServiceOutstandingAtOpen = 0
                    
                    Case mDiff.ServicedCodeVersusServicedExport(Serviced.Comp) _
                     And Not .IsCommon
                        '~~ A non Common Component has  been modified while out of service
                        .Export
                        ServiceOutstandingAtOpen = 0
        
                    Case .CodeExprtd.DiffersFrom(.CodeCrrent) _
                     And .IsCommonUsedOrHosted
                        '~~ A Common Component had been modified while out of service
                        If mDiff.PublicVersusServicedExport(Serviced.Comp) Then
                            '~~ The modification has been made on a (meanwhile) outdated Common Component
                            '~~ The modification will be undone by the update of the outdated version
                            ServiceOutstandingAtOpen = 0
                        Else
                            '~~ The modification had been made based on an up-to-date version
                            If .IsCommCompPending And .IsThePendingSource Then
                                '~~ The serviced Workbook is (still) the source/origin of the pending release
                                .Export
                                CommonPending.Register Serviced.Comp
                                ServiceOutstandingAtOpen = 0
                            Else
                                '~~ The modification is based on an up-to-date version but it conflicts
                                '~~ with another one made meanwhile in another Workbook/VB-Project
                                '~~ This requires a user communication and confirmation
                                ServiceOutstandingAtOpenCase1 Serviced.Comp
                            End If
                        End If
                End Select
            End With
            
        Next vbc
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub ServiceOutstandingAtOpenCase1(ByVal s_comp As clsComp, _
                                 Optional ByVal c_sequ_no As Long = 0)
' ------------------------------------------------------------------------------
' Displays a dialog requesting Confirmed for any used or hosted Common Component
' which has been detected having a service gap.
' ------------------------------------------------------------------------------
    Const PROC = "ServiceOutstandingAtOpens"
    
    On Error GoTo eh
    Dim Comp        As clsComp
    Dim sComp       As String
    Dim v           As Variant
    Dim Msg         As udtMsg
    Dim sBttnConf   As String
    Dim sUsedHosted As String
    Dim sTitle      As String
    
    mBasic.BoP ErrSrc(PROC)
    sBttnConf = "Confirmed"
    With Msg.Section(1)
        .Text.Text = "The concerned " & sUsedHosted & " Common Component's code differs from the current public code in the Common-Components folder. " & _
                     "This modification had yet not been exported which means it must have been modified while the Workbook was " & _
                     "not serviced by CompMan. This concludes to a service gap which cannot be handled other than subsequently " & _
                     "updating the ""outdated"" code. I.e. the made modification will get lost."
    End With
    With Msg.Section(2)
        .Label.Text = mCompMan.BttnAsLabel(sBttnConf)
        .Label.FontColor = rgbBlue
        .Text.Text = "Confirmation is the only choice because it cannot be guaranteed that the modification is based on an up-to-date " & _
                     "code. Displaying the code difference before confirmation may allow to re-do the modification based on an up-to-date " & _
                     "code while the Workbook is serviced."
    End With
    Do
        Select Case mMsg.Dsply(d_title:=sTitle _
                             , d_msg:=Msg _
                             , d_label_spec:="L80" _
                             , d_buttons:=mMsg.Buttons(mDiff.PublicVersusServicedCodeBttn(sComp), vbLf, sBttnConf) _
                             , d_width_min:=350)
            Case mDiff.PublicVersusServicedCodeBttn(sComp): mDiff.PublicVersusServicedCodeDsply Comp
            Case sBttnConf:                                 Exit Do
        End Select
    Loop

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub EnvironmentExportServiceFolderFiles(ByVal e_folder As String)
' ---------------------------------------------------------------------------
' Removes all  Export-Files in the serviced Workbooks's Export-Folder which
' do not correspond with an existing component.
' ---------------------------------------------------------------------------
    Const PROC = "EnvironmentExportServiceFolderFiles"
    
    On Error GoTo eh
    Dim fl                  As File
    Dim sExpFileCorrespComp As String
    Dim sExpFileName        As String
    Dim sPath               As String
    Dim wbk                 As Workbook
    
    With Serviced
        Set wbk = .Wrkbk
        sPath = e_folder
    End With
    
    With fso
        For Each fl In .GetFolder(sPath).Files
            sExpFileCorrespComp = .GetBaseName(fl)
            Select Case .GetExtensionName(fl.Path)
                Case "bas", "cls", "frm", "frx"
                    If Not Serviced.CompExists(sExpFileCorrespComp) Then
                        sExpFileName = .GetFileName(fl.Path)
                        .DeleteFile fl
'                        LogServiced.Entry "Obsolete Export-File """ & sExpFileName & """ deleted (a VBComponent named """ & sExpFileCorrespComp & """ no longer exists)"
                    End If
            End Select
        Next fl
    End With
        
xt: Set fl = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mHskpng." & sProc
End Function

Public Sub FocusOnOpen()
' ------------------------------------------------------------------------------
' Focus on what may have happened concerning any used/hosted Common Component
' between the last CompMan-serviced-export (Workbook_AfterSave event) and this
' Workbook_Open event:
' - Any new public version of a used/hosted Common Component needs to be
'   updated by the corresponding service and thus is not a housekeeping issue.
' - Any yet not registered "pending release" of a used/hosted Common Component
'   indicates a modification which must have been made while the VBProject was
'   not serviced by CompMan. Because it cannot be assured that this modification
'   was based on an up-to-date public version a warning that this modification
'   will be undone by the update service (with an option to display of the
'   difference.
' - All other actions which are also possible between two Workbook_AfterSave
'   events (FocusOnSave)
' ------------------------------------------------------------------------------
    Const PROC = "FocusOnOpen"
    
    On Error GoTo eh
    Dim dctFiles    As Dictionary
    Dim dctComps    As Dictionary
    Dim lTotalItems As Long
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Housekeeping Common Components public
    If CommonPublic Is Nothing Then Set CommonPublic = New clsCommonPublic
    If CommonServiced Is Nothing Then Set CommonServiced = New clsCommonServiced
    Set dctFiles = CommonPublic.ExportFiles
    Set dctComps = CommonPublic.Components
    
    '~~ Calculate number of progress dots
    lTotalItems = dctFiles.Count                                                    ' CommCompsPublicObsolete
    lTotalItems = lTotalItems + dctComps.Count                                      ' CommCompsPublicNew
    lTotalItems = lTotalItems + Serviced.Hosted.Count                               ' CommCompsServiced
    lTotalItems = lTotalItems + CommonServiced.Components.Count                     ' CommCompsServicedNotHosted
    lTotalItems = lTotalItems + Serviced.Wrkbk.VBProject.VBComponents.Count         ' CommCompsServicedKindOf
    lTotalItems = lTotalItems + CommonServiced.Components.Count                     ' CommCompsServicedObsolete
    lTotalItems = lTotalItems + CommonPending.Components.Count                      ' CommCompsServicedPendingRelease
    lTotalItems = lTotalItems + Serviced.Wrkbk.VBProject.VBComponents.Count         ' ServiceOutstandingAtOpens
    lTotalItems = lTotalItems + CommonPublic.Components.Count                       ' CommCompsServicedProperties
    
    With Prgrss
        .Operation = "Housekeeping"
        .ItemsTotal = lTotalItems
    End With
    
    CommCompsPublicObsolete dctFiles, dctComps, 9
    CommCompsPublicNew dctFiles, dctComps, 8
    
    '~~ Housekeeping Common Components serviced
    CommCompsServicedHosted 7
    CommCompsServicedNotHosted 6
    CommCompsServicedKindOf 5
    CommCompsServicedObsolete 4
    
    ServiceOutstandingAtOpen 3
    CommCompsServicedPendingRelease 2
    CommCompsServicedProperties 1
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub FocusOnSave()
' ------------------------------------------------------------------------------
' Manual interactions (to be) covered which may have taken place between two
' Workbook_AfterSave events:
' - Manual copy of a modified Common Component's Export-File to the Common-
'   Components folder
' - Installation/import of a new used Common Component
' - Removal/rename of a used Common Component
' - Common Component no longer hosted but still used
' ------------------------------------------------------------------------------
    Const PROC = "FocusOnSave"
    
    On Error GoTo eh
    Dim dctFiles    As Dictionary
    Dim dctComps    As Dictionary
    Dim lTotalItems As Long
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Housekeeping Common Components public
    If CommonPublic Is Nothing Then Set CommonPublic = New clsCommonPublic
    If CommonServiced Is Nothing Then Set CommonServiced = New clsCommonServiced
    Set dctFiles = CommonPublic.ExportFiles
    Set dctComps = CommonPublic.Components
    
    '~~ Calculate number of progress dots
    lTotalItems = dctFiles.Count                                                    ' CommCompsPublicObsolete
    lTotalItems = lTotalItems + dctComps.Count                                      ' CommCompsPublicNew
    lTotalItems = lTotalItems + Serviced.Hosted.Count                               ' CommCompsServicedHosted
    lTotalItems = lTotalItems + Serviced.Hosted.Count                               ' CommCompsServicedNotHosted
    lTotalItems = lTotalItems + (CommonServiced.Components.Count * 2)               ' CommCompsServiced, CommCompsServicedObsolete
    lTotalItems = lTotalItems + Serviced.Wrkbk.VBProject.VBComponents.Count         ' CommCompsServicedKindOf
    lTotalItems = lTotalItems + (CommonPending.Components.Count * 3)                ' CommCompsServicedPendingRelease
    lTotalItems = lTotalItems + CommonPublic.Components.Count                       ' CommCompsServicedProperties
    
    With Prgrss
        .Operation = "Housekeeping"
        .ItemsTotal = lTotalItems
    End With
    
    CommCompsPublicObsolete dctFiles, dctComps, 8
    CommCompsPublicNew dctFiles, dctComps, 7
    
    '~~ Housekeeping Common Components serviced
    CommCompsServicedHosted 6
    CommCompsServicedNotHosted 5
    CommCompsServicedKindOf 4
    CommCompsServicedObsolete 3
    CommCompsServicedPendingRelease 2
    CommCompsServicedProperties 1
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ProposeContinuationOfModificationInThisWorkbook(ByVal p_comp As String)
' ------------------------------------------------------------------------------
' Dialog providing an option to continue modifications of a Common Component
' started in another Workbook to this Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "ProposeContinuationOfModificationInThisWorkbook"
    
    On Error GoTo eh
    Dim Comp            As clsComp
    Dim Msg             As udtMsg
    Dim sTitle          As String
    Dim sBttnSwitch     As String
    Dim sBttnDoNotSwitch As String
    Dim sRegStatePending As String
    Dim sRegStateSrviced As String
    
    sRegStatePending = mCompMan.RegState(CommonPending.LastModKindOfComp(p_comp))
    sRegStateSrviced = mCompMan.RegState(CommonServiced.KindOfComponent(p_comp))
    
    sBttnDoNotSwitch = "Do not switch with modifications" & vbLf & _
                       "to this Workbook/VBProject" & vbLf & _
                       "I am aware of the concurrent modification risk"
    sBttnSwitch = "Pass on the ongoing modifications to" & vbLf & _
                  "this Workbook. Modifications will be" & vbLf & _
                  "continued and possibly finalized herein."
    
    With Msg.Section(1)
        .Label.Text = "Please note:"
        .Label.FontColor = rgbBlue
        .Text.Text = "The Common Component """ & p_comp & """  " & mBasic.Spaced(sRegStateSrviced) & _
                     "   in this Workbook hase already been modified in Workbook (""" & CommonPending.LastModInWrkbkName(p_comp) & """) " & _
                     "and this modification is yet to finalized/public but still pending release."
    End With
    With Msg.Section(2)
        Select Case True
            Case sRegStatePending = "used" And sRegStateSrviced = "used"
                '~~ Switch from used to used suggested
                sTitle = "Modification switch suggested!"
                .Label.Text = "Recommended:"
                .Label.FontColor = rgbBlue
                .Text.Text = "In order to avoid this pending release is concurrently modified in this Workbook " & _
                             "the continuation is suggested in this Workbook."
            Case sRegStatePending = "used" And sRegStateSrviced = "hosted"
                '~~ Switch from used to hosted suggested
                sTitle = "Modification switch suggested!"
                .Label.Text = "Modification switch strongly recommened"
                .Label.FontColor = rgbBlue
                .Text.Text = "Continuation/finalization of this modification in this current active Workbook is strongly recommended." & vbLf & _
                             "This not only avoids any concurrent modification in this Workbook by accident but also " & _
                             "goes with the modification to the development ""home"" where likely the modification can be tested properly before release to public."
            Case sRegStatePending = "hosted" And sRegStateSrviced = "used"
                '~~ Switch from hosted to used not suggested
                sTitle = "Modification warning!"
                .Label.Text = "Not recommended!"
                .Label.FontColor = rgbRed
                .Label.FontBold = True
                .Text.Text = "In order to avoid an accidential concurrent modification of this pending release " & _
                             "a continuation of the ongoing modification in this active Workbook may be considered. " & _
                             "However, this is   " & mBasic.Spaced("not recommended") & "   because the ongoing modification is " & _
                             "made in the component's ""home"" Workbook, i.e. in the Workbook which claims the component " & _
                             "is ""hosted"", which is supposed to be the Workbook/VBProject which provides all the testing means."

        End Select
    End With
    With Msg.Section(3)
        .Label.Text = "Attention!"
        .Label.FontColor = rgbBlue
        .Text.Text = "When the ongoing modification is not continued in (switched to) this Workbook (because not recommended) " & _
                     "the Common Component  " & mBasic.Spaced("must not be modified") & "   in this Workbook since this would " & _
                     "be considered a conflicting code modification which can not become registered pending release."
        
    End With
    With Msg.Section(4)
        .Label.Text = "Also note"
        .Label.FontColor = rgbBlue
        .Text.Text = "This message will be re-displayed until the pending release modification has been released to public. " & _
                     "Though this is an obvious annoyance, preventing any conflicting code modification is given priority."
    End With
    If mMsg.Dsply(d_title:=sTitle _
                , d_msg:=Msg _
                , d_label_spec:="L100" _
                , d_width_min:=500 _
                , d_buttons:=mMsg.Buttons(sBttnSwitch, vbLf, sBttnDoNotSwitch)) = sBttnSwitch Then
        '~~ Switch ongoing modifications to this Workbook by importing the code
        '~~ from the pennding release Export-File
        mUpdate.ByReImport p_comp, CommonPending.LastModExpFile(p_comp)
        
        Set Comp = New clsComp
        With Comp
            .CompName = p_comp
            .Export
            .SetServicedProperties
            CommonPending.SetPropertiesEqualServiced p_comp
        End With
'        CommonPending.Register Comp
    End If
                         
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub RemoveServiceRemains()
' ------------------------------------------------------------------------------
' 1. Removes any temporary export files created by the update service.
' 2. Removes any components renamed by the update service.
' ------------------------------------------------------------------------------
    Const PROC = "RemoveServiceRemains"
    
    mBasic.BoP ErrSrc(PROC)
    RemoveTempExportFolders
    CommCompsRemoveRenamedByUpdate
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Private Sub RemoveTempExportFolders()
' ------------------------------------------------------------------------------
' Delete a Temp folder possibly created/used by the update service.
' ------------------------------------------------------------------------------
    Dim sFolder As String
    
    sFolder = Servicing.TempExportFolder
    With fso
        If .FolderExists(sFolder) Then .DeleteFolder sFolder
    End With
    
End Sub

