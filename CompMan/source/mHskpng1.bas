Attribute VB_Name = "mHskpng1"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mHskpng: Provides all means for the provision of a consistent
' ======================== and valid data base for all services. Addresses:
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
' PrivProfFileFullName
'
' Public services:
' ----------------
' CommComps                 Maintains an up-to-date data base regarding Common
'                           Components.
' EnvSrvcdWorkbook Maintains and provides the current CompMan services
'                           means environment
' RemoveServiceRemains      Removes any temporary export files created by the
'                           update service and any components renamed by the
'                           update service.
'
' ------------------------------------------------------------------------------
Private Const FILE_NAME_SERVICES_LOG            As String = "Services.log"      ' should never be changed
Private Const FILE_NAME_EXEC_TRACE              As String = "ExecTrace.log"     ' should never be changed
Private Const FILE_NAME_PRIVPROF                As String = "CompMan.dat"
Private Const FOLDER_NAME_COMPMAN_SRVC          As String = "CompMan"

Private sEnvSrvcdCompManServiceFolder           As String
Private sEnvSrvcdExportServiceFolderPath        As String
Private sEnvSrvcdServicesLogFileFullName        As String
Private sEnvSrvcdServicesTraceLogFileFullName   As String
Private sEnvSrvcdPrivProfFileFullName           As String

Public Property Get EnvSrvcdCompManServiceFolder() As String:           EnvSrvcdCompManServiceFolder = sEnvSrvcdCompManServiceFolder:                   End Property

Public Property Get EnvSrvcdExportServiceFolderPath() As String:        EnvSrvcdExportServiceFolderPath = sEnvSrvcdExportServiceFolderPath:             End Property

Public Property Get EnvSrvcdPrivProfFileFullName() As String:           EnvSrvcdPrivProfFileFullName = sEnvSrvcdPrivProfFileFullName:                   End Property

Public Property Get EnvSrvcdServicesLogFileFullName() As String:        EnvSrvcdServicesLogFileFullName = sEnvSrvcdServicesLogFileFullName:             End Property

Public Property Get EnvSrvcdServicesTraceLogFileFullName() As String:   EnvSrvcdServicesTraceLogFileFullName = sEnvSrvcdServicesTraceLogFileFullName:   End Property

Public Sub CommComps()
' ------------------------------------------------------------------------------
' Maintains an up-to-date data base regarding Common Components.
' ------------------------------------------------------------------------------
    Const PROC = "CommComps"
    
    On Error GoTo eh
    Dim dctFiles    As Dictionary
    Dim dctComps    As Dictionary
    Dim lTotalItems As Long
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Housekeeping Common Components public
    If CommonPublic Is Nothing Then Set CommonPublic = New clsCommonPublic
    Set dctFiles = CommonPublic.ExportFiles
    Set dctComps = CommonPublic.Components
    
    lTotalItems = dctFiles.Count                                                    ' CommCompsPublicObsolete
    lTotalItems = lTotalItems + dctComps.Count                                      ' CommCompsPublicNew
    lTotalItems = lTotalItems + Serviced.Hosted.Count                               ' CommCompsServicedHosted
    lTotalItems = lTotalItems + CompManDat.Components.Count                         ' CommCompsServicedNotHosted
    lTotalItems = lTotalItems + Serviced.CompsCommon.Count                          ' CommCompsServicedKindOf
    lTotalItems = lTotalItems + CompManDat.Components.Count                         ' CommCompsServicedObsolete
    If CommonPending.Components.Count <> 0 _
    Then lTotalItems = lTotalItems + Serviced.Wrkbk.VBProject.VBComponents.Count    ' CommCompsServicedPendingRelease
    lTotalItems = lTotalItems + Serviced.CompsCommon.Count                          ' CommCompsServicedPendingReleaseOutstanding
    lTotalItems = lTotalItems + CommonPublic.Components.Count                       ' CommCompsServicedProperties
    
    With Prgrss
        .Operation = "Housekeeping"
        .ItemsTotal = lTotalItems
    End With
    
    CommCompsPublicObsolete dctFiles, dctComps
    If CompManDat.Components.Exists("clsCode") Then Stop
    CommCompsPublicNew dctFiles, dctComps
    If CompManDat.Components.Exists("clsCode") Then Stop
    
    '~~ Housekeeping Common Components serviced
    CommCompsServicedHosted
    If CompManDat.Components.Exists("clsCode") Then Stop
    CommCompsServicedNotHosted
    If CompManDat.Components.Exists("clsCode") Then Stop
    CommCompsServicedKindOf
    If CompManDat.Components.Exists("clsCode") Then Stop
    CommCompsServicedObsolete
    If CompManDat.Components.Exists("clsCode") Then Stop
    CommCompsServicedPendingRelease
    If CompManDat.Components.Exists("clsCode") Then Stop
    CommCompsServicedPendingReleaseOutstanding
    If CompManDat.Components.Exists("clsCode") Then Stop
    CommCompsServicedProperties
    If CompManDat.Components.Exists("clsCode") Then Stop
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function EnvSrvcdCompManServiceFolderCurrentGuessed() As String
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim cll As Collection
    Set cll = mFso.FilesSearch(mCompMan.ServicedWrkbk.Path, FILE_NAME_SERVICES_LOG, 1)
    If cll.Count = 1 Then EnvSrvcdCompManServiceFolderCurrentGuessed = cll(1)
    Set cll = Nothing
    
End Function

Private Sub CommCompsPublicNew(ByVal h_files As Dictionary, _
                                  ByVal h_comps As Dictionary)
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
        For Each v In h_files
            sCompName = v
            sExpFile = h_files(v)
            If Not h_comps.Exists(sCompName) Then
                '~~ The Export-File is yet not registered as a known Common Component
                '~~ It may have been copied manually into the Common-Components folder.
                '~~ I.e. its origin is unknown and thus registered as such.
                '~- The origin will remain unknown until the Common Component is
                '~~ modified in a Workbook using or hosting it and exported.
                With New clsComp
                    .CompName = sCompName
                    If .CodeExprtd.Meets(.CodePublic) Then
                        .SetServicedProperties
                        CompManDat.KindOfComponent(sCompName) = enCompCommonUsed
                        .SetPublicEqualServiced
                    End If
                End With
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

Private Sub CommCompsPublicObsolete(ByVal h_files As Dictionary, _
                                    ByVal h_comps As Dictionary)
' ------------------------------------------------------------------------------
' - Removes in the Common-Component folder's CommComps.dat file any section of
'   which no corresponding Export-File exists in the folder.
' - Removes the corresponding section in the seviced Workbook's CompMan.dat file
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsPublicObsolete"
    
    On Error GoTo eh
    Dim sSection    As String
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Remove any component of which the corresponding Export-File
    '~~ not exists in the Common-Components folder
    For Each v In h_comps
        If Not h_files.Exists(v) Then
            With CommonPublic
                .Remove v
                If .Components.Exists(v) Then .Components.Remove v
                CompManDat.RemoveComponent v
                If CompManDat.Components.Exists("clsCode") Then Stop
            End With
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

Private Sub CommCompsServicedHosted()
' ------------------------------------------------------------------------------
' Maintains in the serviced Workbook's CompMan.dat file for any component
' claimed hosted the corresponding KindOfComponent and the Revision-Number
' (when missing).
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsServicedHosted"
    
    On Error GoTo eh
    Dim wbk         As Workbook
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Serviced.Wrkbk
    
    With CompManDat
        For Each v In Serviced.Hosted
            If Serviced.CompExists(v) Then
                .KindOfComponent(v) = enCompCommonHosted
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

Public Sub CommCompsServicedKindOf()
' ----------------------------------------------------------------------------
' Ensures that all components with a name known as a public Common Component
' are registered either as used or as private. For any yet not registered
' possible public Common Component a dialog is displayed with the two options.
' Note: This procedure is Public testing only.
' ----------------------------------------------------------------------------
    Const PROC = "CommCompsServicedKindOf"
    
    On Error GoTo eh
    Dim BttnPrivate As String
    Dim BttnUsed    As String
    Dim sComp       As String
    Dim dct         As Dictionary
    Dim Msg         As mMsg.udtMsg
    Dim v           As Variant
    Dim vbc         As VBComponent
    Dim sKnownAs    As String
    
    mBasic.BoP ErrSrc(PROC)
    BttnPrivate = "No!" & vbLf & _
                  "This is a  p r i v a t e  component" & vbLf & _
                  "(just by chance with the same name)"
    
    Set dct = CommonPublic.All
    For Each v In Serviced.CompsCommon
        sComp = v
        Select Case CompManDat.KindOfComponent(sComp)
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
                    '~~ The component is a known as a public hosted or pending release Common Component
                    With New clsComp
                        .CompName = sComp
                        Select Case CompManDat.KindOfComponent(sComp)
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
                                
                                Select Case mMsg.Dsply(dsply_title:="Not yet registered ""Common Component""" _
                                                     , dsply_msg:=Msg _
                                                     , dsply_Label_spec:="L30" _
                                                     , dsply_buttons:=mMsg.Buttons(BttnUsed, BttnPrivate))
                                    Case BttnUsed:  CompManDat.KindOfComponent(.CompName) = enCompCommonUsed
                                                    CompManDat.LastModAt(.CompName) = vbNullString ' yet unknown will force update when outdated
                                                    If .IsCommCompUpToDate Then
                                                        .Wrkbk.VBProject.VBComponents(.CompName).Export .ExpFileFullName
                                                        .SetServicedEqualPublic
                                                    End If
                                    Case BttnPrivate:  CompManDat.KindOfComponent(.CompName) = enCompCommonPrivate
                                End Select
                        End Select
                    End With
                End If
                Prgrss.ItemDone = v
        End Select
    Next v

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CommCompsServicedNotHosted()
' ----------------------------------------------------------------------------
' When a former hosting Workbook not or no longer claims a Common Component
' hosted the KindOfComponent in the serviced Workbook's CompMan.dat file
' is changed to enCompCommonUsed.
' ----------------------------------------------------------------------------
    Const PROC      As String = "CommCompsServicedNotHosted"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.ServicedWbk
    For Each v In CompManDat.Components
        If Not Serviced.Hosted.Exists(v) Then
            If mComp.Exists(v, wbk) Then
                CompManDat.KindOfComponent(v) = enCompCommonUsed
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

Private Sub CommCompsServicedObsolete()
'' ------------------------------------------------------------------------------
' Remove sections representing VBComponents no longer existing and those with an
' invalid name.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsServicedObsolete"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    With CompManDat
        For Each v In .Components
            If Not .IsSystemSection(v) Then
                If Not CommonPublic.Exists(v) Or Not Serviced.CompExists(v) Then
                    '~~ The section no longer represents an existing Common Component
                    '~~ or the component no longer exists in the VB-Project
                    .ComponentRemove v
                End If
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

Private Sub CommCompsServicedPendingRelease()
' ------------------------------------------------------------------------------
' Maintains a consistent pending releases data base by:
' 1. Removing pending Common Components of which the Export-file is already
'    identical with the current public Common Component's Export-Fie in the
'    Common-Components folder.
' 2. Removing entries in the Pending.dat file without a corresponding Export-
'    file in the Pending folder
' 3. Removing Export-files without a corresponding entry in the Pending.dat file
'    File in the Common-Components\Pending folder
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsServicedPendingRelease"
    
    On Error GoTo eh
    Dim Comp                        As clsComp
    Dim dctPendingComps             As Dictionary
    Dim dctPendingExpFiles          As Dictionary
    Dim dctPublicComps              As Dictionary
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
    Dim sComp                       As String
    Dim bRemoved                    As Boolean
    Dim vbc                         As VBComponent
    
    mBasic.BoP ErrSrc(PROC)
    '~~ 1. Removing pending Common Components of which the Export-file is identical
    '~~    with the current public Common Component's Export-Fie in the Common-
    '~~    Components folder.
    Set dctPendingComps = CommonPending.Components
    If dctPendingComps.Count <> 0 Then
        Set dctPublicComps = CommonPublic.Components
        For Each vbc In Serviced.Wrkbk.VBProject.VBComponents
            sComp = vbc.Name
            If dctPendingComps.Exists(sComp) Then
                With New clsComp
                    .CompName = sComp
                    If Not .CodePublic.IsNone And Not .CodePnding.IsNone Then
                        '~~ There's already a public Common Component in the Common-Components folder
                        If .CodePublic.Meets(.CodePnding) Then
                            '~~ When the public code is identical with the pending release code the pending is obsolete.
                            '~~ A likely reason is that the export file had been copied manually. The properties
                            '~~ are updated anyway.
                            .SetPublicEqualPending
                            CommonPending.Remove sComp
                        End If
                    End If
                End With
            End If
            Prgrss.ItemDone = v
        Next vbc
    End If
                                            
    With CommonPending
        '~~ 2. Remove Export-files without a corresponding entry in the Pending.dat file
        Set dctPendingExpFiles = .ExportFiles
        For Each v In dctPendingExpFiles
            If Not .Exists(v) Then
                FSo.DeleteFile dctPendingExpFiles(v)
            End If
            Prgrss.ItemDone = v
        Next v
        If dctPendingExpFiles.Count = 0 Then FSo.DeleteFolder CommonPending.FolderPath
    
        '~~ 3. Remove entries in the Pending.dat file without a corresponding Export-
        '~~    File in the Common-Components\Pending folder - may have been moved manually to
        '~~    the Common-Components folder.
        Set dctPendingComps = .Components
        For Each v In dctPendingComps
            sComp = v
            If Not dctPendingExpFiles.Exists(sComp) Then
                With New clsComp
                    .CompName = v
                    .SetPublicEqualPending
                End With
                CommonPending.Remove sComp
            End If
            Prgrss.ItemDone = v
        Next v
    End With
    
    '~~ 4. Regard pending release Common Components of which the the code is
    '~~    identical with the corresponding serviced Workbook's component as
    '~~    "pending release by the serviced Workbook".
    For Each v In CommonPending.Components
        sComp = v
        If Serviced.CompsCommon.Exists(sComp) Then
            Set Comp = New clsComp
            Comp.CompName = sComp
            With Comp
                Select Case True
                    Case .CodePnding.Meets(.CodeCrrent)
                        Serviced.Wrkbk.VBProject.VBComponents(v).Export .ExpFileFullName
                        .SetPendingEqualServiced
                    Case FSo.FileExists(.ExpFileFullName)
                        If .CodePnding.Meets(.CodeExprtd) Then
                            .SetPendingEqualServiced
                        End If
                End Select
            End With
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

Private Sub CommCompsServicedPendingReleaseOutstanding()
' ------------------------------------------------------------------------------
' Regard a serviced Common Component pendig release when:
' - it is not pending release and
' - its code differs from the current public code
' - its LastModAt attribute is greater than the public's LastModAt attribute.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsServicedPendingReleaseOutstanding"
    
    Dim v As Variant
    Dim Comp As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    For Each v In mBasic.KeySort(Serviced.CompsCommon)
        Set Comp = New clsComp
        With Comp
            .CompName = v
            If Not .IsCommCompPending Then
                If .ServicedLastModAt > .PublicLastModAt Then
                    If .CodeCrrent.DiffersFrom(.CodePublic) Then
                        CommonPending.Register Comp
                        .SetPendingEqualServiced
                        
                        With Services
                            .Log(v) = "Serviced Common Component modified (code differs from public and LastModAt > LastmodAt public): Registered pending release"
                            .Log(v) = "Serviced Common Component modified: Properties pending set equal serviced"
                        End With
                    End If
                End If
            End If
        End With
        Prgrss.ItemDone = v
    Next v
    mCompManMenuVBE.Setup

xt: mBasic.EoP ErrSrc(PROC)
End Sub

Private Sub CommCompsServicedProperties()
' ------------------------------------------------------------------------------
' Maintains for all serviced Common Components the properties in the serviced
' workbook's CompMan.dat file - specifically when different from the public
' versions properties differ from the serviced component's properties although
' the code is identical with the public version.
' ------------------------------------------------------------------------------
    Const PROC = "CommCompsServicedProperties"
    
    On Error GoTo eh
    Dim Comp    As clsComp
    Dim sComp   As String
    Dim v       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Loop through all components in the serviced Workbook considered a Commom Component
    For Each v In CommonPublic.Components
        sComp = v
        If sComp = "clsCode" Then Stop
        If Serviced.CompExists(sComp) Then
            With New clsComp
                .CompName = sComp
                '~~ For the Common Component exists a public version in the Common-Components folder
                If Not .ServicedMeetPublicProperties Then
                    '~~ Just in case the "local" serviced component's properties
                    '~~ in the CompMan.dat file differ from those public
                    If .CodeCrrent.Meets(.CodePublic) Then
                        '~~ The public version (Export-File in th Common-Components folder) has obviously been manually imported
                        .SetServicedEqualPublic
                        If Not FSo.FileExists(.ExpFileFullName) Then
                            '~~ Apparently the public Common Component's Export-File has manually been imported
                            Serviced.Wrkbk.VBProject.VBComponents(v).Export .ExpFileFullName
                            With Services
                                .Log(sComp) = "Serviced Common Component properties housekeeping:"
                                .Log(sComp) = "Serviced Common Component (public assumed imported manually): e x p o r t e d !"
                                .Log(sComp) = "Serviced Common Component (public assumed imported manually): Properties set equal public"
                            End With
                        End If
                    ElseIf FSo.FileExists(.ExpFileFullName) Then
                        '~~ The current code may have already been modified
                        If .CodeExprtd.Meets(.CodePublic) Then
                            .SetServicedEqualPublic
                            With Services
                                .Log(sComp) = "Serviced Common Component properties housekeeping:"
                                .Log(sComp) = "Serviced Common Component's Export-File meets public: Properties set equal public"
                            End With
                        End If
                    End If
                End If
            End With
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

Public Sub EnvSrvcd()
' ------------------------------------------------------------------------------
' Maintains and provides the current CompMan services means environment for
' a serviced Workbook aiming for a single folder which contains everything
' CompMan requires for the provi.
' Item/structure                            Current name/location
' ----------------------------------------- ------------------------------------
' <service-folder>                          <compman-root>\CompMan
'  +- <service-trace-log-file>              <compman-root>\CompMan\ExecTrace.log
'  +- <commm-comps-private-profile-file>    <compman-root>\CompMan\CompMan.dat
'  +- <export-folder>                       <compman-root>\source
'  +- <service-log-file>                    <compman-root>\CompMan\Service.log
' ------------------------------------------------------------------------------
    Const PROC          As String = "EnvSrvcdWorkbook"
    Dim sFolder         As String
    Dim sCurrentRoot    As String
    Dim sCurrentGuessed As String
    
    On Error GoTo eh
    
    sCurrentGuessed = FSo.GetFile(EnvSrvcdCompManServiceFolderCurrentGuessed).ParentFolder.Path
    
    '~~ <service-folder>
    sEnvSrvcdCompManServiceFolder = mCompMan.ServicedWrkbk.Path & "\" & FOLDER_NAME_COMPMAN_SRVC
    If sCurrentGuessed = vbNullString Then
        '~~ No current service folder found
        If Not FSo.FolderExists(sEnvSrvcdCompManServiceFolder) Then FSo.CreateFolder sEnvSrvcdCompManServiceFolder
    ElseIf sCurrentGuessed <> sEnvSrvcdCompManServiceFolder Then
        FSo.GetFolder(sCurrentGuessed).Name = FOLDER_NAME_COMPMAN_SRVC
    End If
    
    '~~ <service-trace-log-file>
    sEnvSrvcdServicesTraceLogFileFullName = sEnvSrvcdCompManServiceFolder & "\" & FILE_NAME_EXEC_TRACE
    EnvSrvcdEstablishExecTraceLog

    '~~ <commm-comps-private-profile-file>
    sEnvSrvcdPrivProfFileFullName = sEnvSrvcdCompManServiceFolder & "\" & FILE_NAME_PRIVPROF
    Set CompManDat = New clsCompManDat
    
    '~~ <export-folder>
    sEnvSrvcdExportServiceFolderPath = EnvSrvcdExportServiceFolderName
    
    '~~ <service-log-file>
    sEnvSrvcdServicesLogFileFullName = sEnvSrvcdCompManServiceFolder & "\" & FILE_NAME_SERVICES_LOG
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub EnvSrvcdEstablishExecTraceLog()
' --------------------------------------------------------------------------
' - Establishes the execution trace log file for the serviced Workbook in
'   the serviced Workbook's CompMan.ServicesFolder.
' - Considers the mTrc Standard Module is used (Cond. Comp. Arg. `mTrc = 1`)
'   or the clsTrc class module is used (Cond. Comp. Arg. `clsTrc = 1`)
' - Establishes a new log file along with the update service only (else the
'   log is appended).
' --------------------------------------------------------------------------

#If mTrc = 1 Then
    mTrc.FileFullName = Hskpng.ServicesTraceLogFileFullName
    mTrc.KeepLogs = 5
    mTrc.Title = Services.CurrentServiceStatusBar
    If sCurrentServiceName = mCompManClient.SRVC_UPDATE_OUTDATED_DSPLY _
    Then mTrc.NewFile
    sExecTraceFile = mTrc.FileFullName

#ElseIf clsTrc = 1 Then
    Set Trc = Nothing: Set Trc = New clsTrc
    With Trc
        .FileFullName = sEnvSrvcdServicesTraceLogFileFullName
        .KeepLogs = 5
        .Title = mCompMan.CurrentServiceStatusBar
        If mCompMan.CurrentServiceName = mCompManClient.SRVC_UPDATE_OUTDATED_DSPLY Then
            .NewFile
        End If
    End With
#End If

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "Hskpng." & sProc
End Function

Private Sub EnvSrvcdExportServiceFolderFiles(ByVal e_folder As String)
' ---------------------------------------------------------------------------
' Removes all  Export-Files in the serviced Workbooks's Export-Folder which
' do not correspond with an existing component.
' ---------------------------------------------------------------------------
    Const PROC = "EnvSrvcdExportServiceFolderFiles"
    
    On Error GoTo eh
    Dim fl                  As File
    Dim sExpFileCorrespComp As String
    Dim sExpFileExtension   As String
    Dim sExpFileName        As String
    Dim sPath               As String
    Dim wbk                 As Workbook
    
    With Serviced
        Set wbk = .Wrkbk
        sPath = e_folder
    End With
    
    With FSo
        For Each fl In .GetFolder(sPath).Files
            sExpFileCorrespComp = .GetBaseName(fl)
            Select Case .GetExtensionName(fl.Path)
                Case "bas", "cls", "frm", "frx"
                    If Not mComp.Exists(sExpFileCorrespComp, wbk) Then
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

Private Function EnvSrvcdExportServiceFolderCurrentGuessed(ByVal e_recent_name As String) As String
' ---------------------------------------------------------------------------
' Returns the path for the recently used export folder
' ---------------------------------------------------------------------------
    Dim cll As Collection
    Dim v   As Variant
    
    Set cll = mFso.Folders(mCompMan.ServicedWrkbk.Path, True)
    For Each v In cll
        If InStr(v, "\" & e_recent_name) <> 0 _
        And InStr(v, "\Addin") = 0 Then
            EnvSrvcdExportServiceFolderCurrentGuessed = v
            Exit For
        End If
    Next v
    Set cll = Nothing
    
End Function

Private Function EnvSrvcdExportServiceFolderName() As String
' ---------------------------------------------------------------------------
' Create a missing Export-Folder for the serviced Workbook or forward an
' outdated (last used) existing Export-Folder.
' ---------------------------------------------------------------------------
    Const PROC = "EnvSrvcdExportServiceFolderName"
    
    On Error GoTo eh
    Dim sExpFldrNameCurrent As String
    Dim sExpFldrNameRecent  As String
    Dim sExpFldrPrntCurrent As String
    Dim sExpFldrPrntRecent  As String
    
    sExpFldrNameRecent = CompManDat.RecentlyUsedExportFolder
    sExpFldrPrntRecent = FSo.GetFolder(EnvSrvcdExportServiceFolderCurrentGuessed(sExpFldrNameRecent)).ParentFolder.Path
    
    sExpFldrNameCurrent = wsConfig.FolderExport ' note that configured is only the name but not the location
    sExpFldrPrntCurrent = Hskpng.EnvSrvcdCompManServiceFolder
    
    If sExpFldrPrntCurrent <> sExpFldrPrntRecent _
    Or sExpFldrNameRecent <> sExpFldrNameCurrent Then
    
        Select Case True
        '~~ 1. Make sure its name is up-to-date
        '~~ The export folder's name is outdated and/or its location
            Case sExpFldrPrntRecent = sExpFldrPrntCurrent And sExpFldrNameRecent = sExpFldrNameCurrent
                '~~ Correct name and location
                Stop
            Case sExpFldrPrntRecent = sExpFldrPrntCurrent And sExpFldrNameRecent <> sExpFldrNameCurrent
                '~~ The Export folder resides in its correct parent folder but with an outdated name
                FSo.GetFolder(sExpFldrPrntCurrent & "\" & sExpFldrNameRecent).Name = sExpFldrNameCurrent
                CompManDat.RecentlyUsedExportFolder = sExpFldrNameCurrent
            
            Case sExpFldrPrntRecent <> sExpFldrPrntCurrent And sExpFldrNameRecent = sExpFldrNameCurrent
                '~~ The Export folder has its correct name but resides at the wrong location
                FSo.MoveFolder sExpFldrPrntRecent & "\" & sExpFldrNameRecent, sExpFldrPrntCurrent & "\" & sExpFldrNameRecent
                
            Case sExpFldrPrntRecent <> sExpFldrPrntCurrent And sExpFldrNameRecent <> sExpFldrNameCurrent
                FSo.MoveFolder sExpFldrPrntRecent & "\" & sExpFldrNameRecent, sExpFldrPrntCurrent
                FSo.GetFolder(sExpFldrPrntCurrent & "\" & sExpFldrNameRecent).Name = sExpFldrNameCurrent
                CompManDat.RecentlyUsedExportFolder = sExpFldrNameCurrent
                
        End Select
    End If
                
xt: EnvSrvcdExportServiceFolderName = sExpFldrPrntCurrent & "\" & sExpFldrNameCurrent
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub RemoveRenamedByUpdate()
' ------------------------------------------------------------------------------
' Removes all Common Components which had been renamed in order to enable the
' import of an Export-File (from the Common-Components folder) wich represents
' an up-to-date code version of a Common Component.
' ------------------------------------------------------------------------------
    Const PROC = "RemoveRenamedByUpdate"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim sComp   As String
    
    With Serviced.Wrkbk.VBProject
        For Each vbc In .VBComponents
            sComp = vbc.Name
            If mFact.HasBeenRenamedByUpdateService(sComp) Then
                Services.Log(sComp) = "Component """ & sComp & """ (resulting from rename to enable import of up-to-date Common Component) removed."
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

Public Sub RemoveServiceRemains()
' ------------------------------------------------------------------------------
' 1. Removes any temporary export files created by the update service.
' 2. Removes any components renamed by the update service.
' ------------------------------------------------------------------------------
    Const PROC = "RemoveServiceRemains"
    
    mBasic.BoP ErrSrc(PROC)
    RemoveTempExportFolders
    RemoveRenamedByUpdate
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Private Sub RemoveTempExportFolders()
' ------------------------------------------------------------------------------
' Delete a Temp folder possibly created/used by the update service.
' ------------------------------------------------------------------------------
    Dim sFolder As String
    
    sFolder = Services.TempExportFolder
    With FSo
        If .FolderExists(sFolder) Then .DeleteFolder sFolder
    End With
    
End Sub

