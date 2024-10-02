VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Class Module clsCommComps: Management of Common Components
' ========================== a) in the dedicated Common-Components folder
'                            b) in the CommComps.dat file
' Public Methods:
' ----------------
' BttnInconsistencyExport          .
' BttnInconsistencySkip            .
' BttnInconsistencyUpdate          .
' CommCompsDatFileFullName         .
'
' Public Properties:
' ------------------
' CompName                 w Essential throughout all services/methods
' Hosted                   w Provides the hosted components, provided as a
'                            comma delimited string in a Dictionary
'                            (dctHosted) and performs the housekeeping for
'                            the Common-Components folder's CommComps.dat
'                            file
' LastMod...             r/w Properties are provided whenever a used/hosted
'                            Common Component's code has changed, along with
'                            the clsComp.ReleaseComp method.
' PendingReleaseRegister   w Modifications made for a used/hosted Common
'                            Component are registered 'pending release'.
'
' ---------------------------------------------------------------------------

' Attention: To change any value names check syntax
Private Const EXP_FILE_EXTENTION                        As String = "ExportFileExtention"
Private Const LAST_MODIFIED_AT_DATE_TIME                As String = "Last_Modified_AtDateTime"
Private Const LAST_MODIFIED_EXP_FILE_FULL_NAME_ORIGIN   As String = "LastModExpFileFullNameOrigin"
Private Const LAST_MODIFIED_IN_WBK_FULL_NAME            As String = "Last_Modified_InWbkFullName"
Private Const LAST_MODIFIED_IN_WBK_NAME                 As String = "Last_Modified_InWbkName"
Private Const LAST_MODIFIED_ON_MACHINE                  As String = "Last_Modified_OnMachine"

Private PPCommComps                                     As clsPP
Private dctHosted                                       As Dictionary
Private sFolderPath                                     As String
Private sCompName                                       As String
Private sCommCompsDatFile                               As String
Private PP                                              As New clsPrivProf

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
    
    Dim dct         As New Dictionary
    Dim sHskpngNms  As String
    
    mBasic.BoP ErrSrc(PROC)
'    Set dctHosted = New Dictionary
    Set PPCommComps = New clsPP
    
    sFolderPath = wsConfig.FolderCommonComponentsPath

    '~~> Full name of the PPCommComps-File
    sCommCompsDatFile = sFolderPath & "\" & "CommComps.dat"
    
    PPCommComps.FileFullName = sCommCompsDatFile
    PP.FileName = sCommCompsDatFile
    
    ' Housekeeping names syntax:
    ' [[sect:][sect:]][current-name]<prev-name<prev-name
    ' [[sect:][sect:]]prev-name>prev-name>[current-name]
    ' Rule 1: > or < indicate the history sequence
    ' Rule 2: One ot more sections are optional. When specified, the names history
    '         applies only for those, when omitted for all sections the name any
    '         of the names in the history is used.
    ' Rule 3: The > and the < character point to the predecessor of a name. I.e.
    '         both syntaxes are valid.
    ' Rule 4: The current name is optional. I.e. when > or < is not followed by a
    '         name (the current new name), the life any history name has ended.
    '         I. e. the name is no longer used at all and thus any in use history
    '         name is removed - in the specified or all sections.
    
    PP.HskpngNames PP.FileName, _
                   "RawRevisionNumber>" _
                 , LAST_MODIFIED_AT_DATE_TIME & "<LastModifiedAtDateTime<RevisionNumber" _
                 , LAST_MODIFIED_EXP_FILE_FULL_NAME_ORIGIN & "<LastModExpFileFullName" _
                 , LAST_MODIFIED_IN_WBK_FULL_NAME & "<LastModifiedInWbkFullName<LastModWbkFullName" _
                 , LAST_MODIFIED_IN_WBK_NAME & "<LastModifiedInWbkFullName<LastModWbkFullName" _
                 , LAST_MODIFIED_ON_MACHINE & "<LastModifiedOnMachine"
    mBasic.EoP ErrSrc(PROC)

End Sub

Public Sub DsplyDiffFormerModificationsVersusCurrentPublic(ByVal d_comp As clsComp)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
                    
    With d_comp
        .CodeCurrent.Source = .ExpFileFullName
        .CodeCurrent.DsplyDiffs d_this_file_name:="ModificationsInThisWorkbook" _
                                 , d_this_file_title:="Current exported last modifications of the Common Component  " & mBasic.Spaced(sCompName) _
                                 , d_from_code:=.CodePublic _
                                 , d_from_file_name:="CurrentPublicCommonComponent" _
                                 , d_from_file_title:="Current public code of Common Component  " & mBasic.Spaced(sCompName)
    End With
    
End Sub

Private Sub Class_Terminate()
    Set PPCommComps = Nothing
End Sub

Friend Property Get BttnDsplyDiff() As String: BttnDsplyDiff = "Display difference" & vbLf & "(of Export-Files)":   End Property

Public Property Let CompName(ByVal c_name As String): sCompName = c_name: End Property

Friend Property Let Hosted(ByVal s As String)
' ---------------------------------------------------------------------------
' Provides the hosted components, provided as a comma delimited string in a
' Dictionary (dctHosted) and performs the housekeeping for Common Components
' in the Common-Components folder's CommComps.dat file.
' ---------------------------------------------------------------------------
    Dim v       As Variant
    Dim sComp   As String
    Dim dct     As New Dictionary
    
    Set dctHosted = Nothing: Set dctHosted = New Dictionary
    For Each v In Split(s, ",")
        sComp = Trim$(v)
        If Not dct.Exists(sComp) Then
            dctHosted.Add sComp, sComp
        End If
    Next v
    
End Property

Friend Property Get LastModAtDateTimeUTC(Optional ByVal l_comp_name As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModAtDateTimeUTC = Value(LAST_MODIFIED_AT_DATE_TIME, l_comp_name)
End Property

Friend Property Let ExportFileExtention(Optional ByVal l_comp_name As String = vbNullString, _
                                                 ByVal l_ext As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(EXP_FILE_EXTENTION, l_comp_name) = l_ext
End Property
                                                 
Friend Property Get ExportFileExtention(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    ExportFileExtention = Value(EXP_FILE_EXTENTION, l_comp_name)
End Property

Friend Property Let LastModAtDateTimeUTC(Optional ByVal l_comp_name As String = vbNullString, _
                                               ByVal l_comp_rev_no As String)
' ------------------------------------------------------------------------------
' Returns a revision number in the form yy-mm-dd.00. Plus one when an existing
' revision number is provided (comp_rev-no) or the current date with .01.
' ------------------------------------------------------------------------------
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    If l_comp_rev_no = vbNullString Then l_comp_rev_no = CompMan.UTC & " (UTC)"
    Value(LAST_MODIFIED_AT_DATE_TIME, l_comp_name) = l_comp_rev_no
End Property

Friend Property Get LastModExpFile(Optional ByVal l_comp_name As String = vbNullString) As File
' ---------------------------------------------------------------------------
' Returns a component's (comp_name) Export-File in the Common-Components
' folder as file object.
' ---------------------------------------------------------------------------
    Const PROC = "LastModExpFile Get"
    
    On Error GoTo eh
    Dim fl As File
    
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    With FSo
        For Each fl In .GetFolder(sFolderPath).Files
            If .GetBaseName(fl.Path) = l_comp_name Then
                Set LastModExpFile = fl
                Exit For
            End If
        Next fl
    End With
    
xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let LastModExpFile(Optional ByVal l_comp_name As String = vbNullString, _
                                             ByVal l_comp_exp_file As File)
' ---------------------------------------------------------------------------
' Copies the Raw Export File from its host Workbook location to the Common
' Components Folder from where it is used as the source for the update of
' outdated Used Common Components.
' ---------------------------------------------------------------------------
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    l_comp_exp_file.Copy Destination:=sFolderPath & l_comp_exp_file.Name, OverWriteFiles:=True
    Me.LastModExpFileFullNameOrigin(l_comp_name) = l_comp_exp_file.Path
End Property

Friend Property Get LastModExpFileFullName(Optional ByVal l_comp_name As String = vbNullString) As String
' ---------------------------------------------------------------------------
' Returns the Common-Component folder's Export-File Full Name.
' ---------------------------------------------------------------------------
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    With FSo
        LastModExpFileFullName = sFolderPath & "\" & l_comp_name & "." & ExportFileExtention(l_comp_name)
    End With
End Property

Friend Property Get LastModExpFileFullNameOrigin(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModExpFileFullNameOrigin = Value(LAST_MODIFIED_EXP_FILE_FULL_NAME_ORIGIN, l_comp_name)
End Property

Friend Property Let LastModExpFileFullNameOrigin(Optional ByVal l_comp_name As String = vbNullString, _
                                                          ByVal l_exp_file_full_name As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(LAST_MODIFIED_EXP_FILE_FULL_NAME_ORIGIN, l_comp_name) = l_exp_file_full_name
End Property

Friend Property Let LastModInWbk(Optional ByVal l_comp_name As String = vbNullString, _
                                          ByVal l_wbk As Workbook)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    If Not l_wbk Is Nothing Then
        Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name) = l_wbk.FullName
        Value(LAST_MODIFIED_IN_WBK_NAME, l_comp_name) = l_wbk.Name
    Else
        Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name) = vbNullString
        Value(LAST_MODIFIED_IN_WBK_NAME, l_comp_name) = vbNullString
    End If
End Property

Friend Property Get LastModInWbkBaseName(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModInWbkBaseName = FSo.GetBaseName(Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name))
End Property

Friend Property Get LastModInWbkFullName(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModInWbkFullName = Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name)
End Property

Friend Property Let LastModInWbkFullName(Optional ByVal l_comp_name As String = vbNullString, _
                                                  ByVal l_full_name As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name) = l_full_name
End Property

Friend Property Get LastModInWbkName(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModInWbkName = Value(LAST_MODIFIED_IN_WBK_NAME, l_comp_name)
    If LastModInWbkName = vbNullString Then
        LastModInWbkName = FSo.GetFileName(Value(LAST_MODIFIED_IN_WBK_FULL_NAME, l_comp_name))
    End If
End Property

Friend Property Let LastModInWbkName(Optional ByVal l_comp_name As String = vbNullString, _
                                              ByVal l_wbk_name As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(LAST_MODIFIED_IN_WBK_NAME, l_comp_name) = l_wbk_name
End Property
Friend Property Get LastModOnMachine(Optional ByVal l_comp_name As String = vbNullString) As String
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    LastModOnMachine = Value(LAST_MODIFIED_ON_MACHINE, l_comp_name)
End Property

Friend Property Let LastModOnMachine(Optional ByVal l_comp_name As String = vbNullString, _
                                              ByVal l_machine_name As String)
    If l_comp_name = vbNullString Then l_comp_name = sCompName
    Value(LAST_MODIFIED_ON_MACHINE, l_comp_name) = l_machine_name
End Property

Private Property Get Value(Optional ByVal v_value_name As String = vbNullString, _
                           Optional ByVal v_comp_name As String = vbNullString) As String
    Const PROC = "Value/Get"
    
    On Error GoTo eh
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No value-name provided!"
    
    If v_comp_name = vbNullString Then
        If sCompName = vbNullString Then
            Err.Raise AppErr(2), ErrSrc(PROC), "No component name provided, section name is missing!"
        Else
            v_comp_name = sCompName
        End If
    End If
    With PPCommComps
        Value = .Value(pp_section:=v_comp_name, pp_value_name:=v_value_name)
    End With

xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let Value(Optional ByVal v_value_name As String, _
                           Optional ByVal v_comp_name As String = vbNullString, _
                                    ByVal v_value As String)
    Const PROC = "Value/Let"
    
    On Error GoTo eh
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No value-name provided!"
    
    If v_comp_name = vbNullString Then
        If sCompName = vbNullString Then
            Err.Raise AppErr(1), ErrSrc(PROC), "No component name provided, section name is missing!"
        Else
            v_comp_name = sCompName
        End If
    End If
    With PPCommComps
        .Value(pp_section:=v_comp_name, pp_value_name:=v_value_name) = v_value
    End With
                                    
xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Sub ComponentRemove(ByVal s As String)
     PPCommComps.ComponentRemove s
End Sub

Public Function Components() As Dictionary
    Const PROC = "Components"
    
    mBasic.BoP ErrSrc(PROC)
    Set Components = PP.Sections
    mBasic.EoP ErrSrc(PROC)
    
End Function

Public Property Get CurrentPublicExpFileFullName() As String
    With FSo
        Select Case True
            Case .FileExists(sFolderPath & "\" & sCompName & ".bas"): CurrentPublicExpFileFullName = sFolderPath & "\" & sCompName & ".bas"
            Case .FileExists(sFolderPath & "\" & sCompName & ".cls"): CurrentPublicExpFileFullName = sFolderPath & "\" & sCompName & ".cls"
            Case .FileExists(sFolderPath & "\" & sCompName & ".frm"): CurrentPublicExpFileFullName = sFolderPath & "\" & sCompName & ".frm"
        End Select
    End With
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCommComps" & "." & sProc
End Function

Public Function ExistsPublic(ByVal e_comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Common-Components folder has a an Export-File for
' the named component (e_comp_name), i.e. there is a public code version
' available.
' ----------------------------------------------------------------------------
    ExistsPublic = LastModInWbkFullName(e_comp_name) <> vbNullString
End Function

Public Function ExistsRegistered(ByVal x_comp_name As String) As Boolean
    ExistsRegistered = Components.Exists(x_comp_name)
End Function

Public Sub Hskpng()
' ------------------------------------------------------------------------------
' Removes obsolete sections which are those neither representing an existing
' VBComponent no another valid section's Name.
' ------------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    Dim dct As Dictionary
    Dim wbk As Workbook
    Dim fld As Folder
    Dim fls As New Dictionary
    Dim fl  As File
    Dim s   As String
    
    mBasic.BoP ErrSrc(PROC)
    If Services Is Nothing Then
        Set Services = New clsServices
        Services.Initiate i_service_proc:=mCompManClient.SRVC_RELEASE_PENDING _
                        , i_serviced_wbk:=ThisWorkbook
    End If
    
    '~~ Collect the base-name of all files in the folder each representing
    '~~ an available public Common Component's code
    Set fld = FSo.GetFolder(sFolderPath)
    With fls
        For Each fl In fld.Files
            s = FSo.GetBaseName(fl.Path)
            If Not .Exists(s) _
            Then .Add s, vbNullString
        Next fl
    End With

    Set wbk = Services.ServicedWbk
    Set dct = PP.Sections
    
    Set Prgrss = New clsProgress
    With Prgrss
        .Operation = "Hskpng"
        .ItemsTotal = dct.Count _
                    + fld.Files.Count _
                    + wbk.VBProject.VBComponents.Count _
                    + dctHosted.Count
        .Figures = False
        .DoneItemsInfo = False
    End With
    HskpngSectionsObsolete dct, fls
    HskpngSectionsMissing fld, dct
    HskpngUsedOrPrivate wbk
    HskpngHosted
    Reorg
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngHosted()
' ----------------------------------------------------------------------------
' Maintains for each Commmon Component claimed hosted by the serviced Workbook
' the following values in the Common-Components folder's CommComp.dat file
' when yet not provided:
' - Component Name (section)
' - LastModifiedInWbkFullName (value)
' - Revision Number
' ----------------------------------------------------------------------------
    Const PROC = "HskpngHosted"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim sHostBaseName   As String
    Dim wbk             As Workbook
    Dim Comp            As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.ServicedWbk
    sHostBaseName = FSo.GetBaseName(wbk.FullName)
    
    For Each v In dctHosted
        If Not mComp.Exists(v, wbk) Then
            MsgBox "The VBComponent " & v & " claimed hosted by the serviced Workbook " & Services.ServicedWbk.Name & _
                   " will be ignored because it does not exist in the Workbook/VB-Project." & vbLf & vbLf & _
                   "The hosted argument should be updated accordingly.", _
                   vbOK, "VBComponent " & v & "does not exist!"
            Prgrss.ItemSkipped
        Else
            '~~ When the Common Component claimed hosted is not identical with the current public code
            '~~ or it has yet never been released public it will here be registered for being released.
            Set Comp = New clsComp
            With Comp
                .Wrkbk = wbk
                .CompName = v
            End With
            Pending.Management Comp
            Prgrss.ItemDone = Comp.CompName
            Set Comp = Nothing
        End If
    Next v

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngSectionsMissing(ByVal h_fld As Folder, _
                                  ByVal h_comps As Dictionary)
' ------------------------------------------------------------------------------
' Adds for any Export-File in the Common-Components folder a section in the
' Common-Component folder's CommComps.dat file a section.
'
' Background:
' A missing section indicates a Common Component of wich the Export-File has
' obvously been copied manually into the Common-Components folder which now in
' the sense of CompMan has become an available Common Component ready for being
' imported into any VB-Project. A new registered Common Component remains
' un-hosted until a Workbbok claims hosting it, i.e. providing a delevelopment
' and test environment for it.
' ------------------------------------------------------------------------------
    Const PROC = "HskpngSectionsMissing"
    
    On Error GoTo eh
    Dim fle         As File
    Dim sCompName   As String
    Dim sExt        As String
    
    mBasic.BoP ErrSrc(PROC)
    With FSo
        For Each fle In h_fld.Files
            sExt = .GetExtensionName(fle.Path)
            Select Case sExt
                Case "bas", "frm", "cls"
                    sCompName = .GetBaseName(fle.Path)
                    If Not h_comps.Exists(sCompName) Then
                        '~~ The Export-File is yet not registered as a known Common Component
                        '~~ It most likely has been copied manually into the Common-Components
                        '~~ folder. I.e. its "raw host" is unknown - and registered as such.
                        '~- The raw host will remain unknown until the Common Component is
                        '~~ modified in a Workbook using it and exported.
                        LastModInWbk(sCompName) = Nothing
                        LastModAtDateTimeUTC(sCompName) = CompMan.UTC & " (UTC)"
                        Prgrss.ItemDone = sCompName
                    Else
                        Prgrss.ItemSkipped
                        If LastModExpFileFullNameOrigin(sCompName) = vbNullString Then
'                            Debug.Print "The property ""LastModExpFileFullNameOrigin"" of component " & sCompName & " is not available, i.e. its origin is unknown or simply yet not registered respectively!"
                        End If
                    End If
                Case Else
                    Prgrss.ItemSkipped
            End Select
        Next fle
    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngSectionsObsolete(ByVal h_comps As Dictionary, _
                                   ByVal h_files As Dictionary)
' ------------------------------------------------------------------------------
' Removes in the Common-Component folder's CommComps.dat file any section of
' which no corresponding Export-File exists in the folder
' ------------------------------------------------------------------------------
    Const PROC = "HskpngSectionsObsolete"
    
    On Error GoTo eh
    Dim sSection    As String
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
        
    '~~ Remove any component of which the corresponding Export-File
    '~~ not exists in the Common-Components folder
    For Each v In h_comps
        sSection = v
        If sSection <> PP.NameHskpngSection Then ' management section does not have a corresponding file
            If Not h_files.Exists(v) Then
                PP.SectionRemove sSection
                Prgrss.ItemDone = sSection
            Else
                Prgrss.ItemSkipped
            End If
        Else
            Prgrss.ItemSkipped
        End If
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngUsedOrPrivate(ByVal h_wbk As Workbook)
' ----------------------------------------------------------------------------
' Manages the registration of used Common Components, done before change
' components are exported and used changed Common Components are updated.
' When not yet registered a confirmation dialog ensures a component not just
' accidentially has the same name. The type of confirmation is registered
' either as "used" or "private" together with the current revision number.
' When none is available the current date is registered on the fly.
' its Revision-Number is
' ----------------------------------------------------------------------------
    Const PROC = "HskpngUsedOrPrivate"
    
    On Error GoTo eh
    Dim vbc                     As VBComponent
    Dim BttnConfirmedUsed       As String
    Dim BttnConfirmedPrivate    As String
    Dim Msg                     As mMsg.udtMsg
    Dim Comp                    As clsComp
    Dim lRegState               As Long
    
    mBasic.BoP ErrSrc(PROC)
    BttnConfirmedUsed = "Yes!" & vbLf & _
                        "This is a used Common Component" & vbLf & _
                        "identical with the corresponding" & vbLf & _
                        "VBComponent's Export-File in the" & vbLf & _
                        """Common-Components folder"""
    BttnConfirmedPrivate = "No!" & vbLf & _
                           "This is a VBComponent which just" & vbLf & _
                           "by chance has the same name."
    
    For Each vbc In h_wbk.VBProject.VBComponents
        If Not ExistsPublic(vbc.Name) Then
            Prgrss.ItemSkipped
        Else
            Set Comp = New clsComp
            With Comp
                .Wrkbk = h_wbk
                .VBComp = vbc
                '~~ The component is a known Common Component
                Select Case PPCompManDat.RegistrationState(.CompName)
                    Case enRegStatePrivate: Prgrss.ItemSkipped
                    Case enRegStateUsed:    Prgrss.ItemSkipped
                    Case enRegStateHosted:  Prgrss.ItemSkipped
                    Case Else
                        '~~ Once an equally named VBComponent is registered as private it will no longer be regarded as "used" and updated.
                        Msg.Section(1).Text.Text = "The component in the VBProject named   " & mBasic.Spaced(.CompName) & "   is known as a ""Common Component"", " & _
                                                   "i.e. an equally named component exists in the ""Common-Components folder""  '" & _
                                                   sFolderPath & "', but the component is yet neither registered/known as ""used"" nor as ""private"" !" & vbLf & _
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
                                             , dsply_buttons:=mMsg.Buttons(BttnConfirmedUsed, BttnConfirmedPrivate))
                            Case BttnConfirmedUsed:     PPCompManDat.RegistrationState(.CompName) = enRegStateUsed
                                                        PPCompManDat.LastModAtDateTimeUTC(.CompName) = vbNullString ' yet unknown will force update when outdated
                            Case BttnConfirmedPrivate:  PPCompManDat.RegistrationState(.CompName) = enRegStatePrivate
                        End Select
                        Prgrss.ItemDone = .CompName
                End Select
            End With
            Set Comp = Nothing
        End If
    Next vbc

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function LastModExpFileExists(ByVal comp_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when an Export File with the name extracted from the
' LastModExpFileFullNameOrigin exists in the Common Components Folder
' ----------------------------------------------------------------------------
    Dim FileName    As String
    With New FileSystemObject
        FileName = .GetFileName(LastModExpFileFullNameOrigin(comp_name))
        LastModExpFileExists = .FileExists(sFolderPath & "\" & FileName)
    End With
End Function

Private Function MaxRawLenght() As Long
' -----------------------------------------------
' Returns the max length of a raw componen's name
' -----------------------------------------------
    Const PROC = "MaxRawLenght"
    
    On Error GoTo eh
    Dim v As Variant
    
    For Each v In Components
        MaxRawLenght = Max(MaxRawLenght, Len(v))
    Next v
    
xt: Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Reorg()
    Const PROC = "Reorg"
    mBasic.BoP ErrSrc(PROC)
    PPCommComps.Reorg
    mBasic.EoP ErrSrc(PROC)
End Sub



