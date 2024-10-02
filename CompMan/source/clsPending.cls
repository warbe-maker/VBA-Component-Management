VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------------------------------------------------------
' Class Module clsPending: Fokus on modified, and thus pending release, Common
' ======================== Components. This pending management uses a Pending folder
' in Common Componets folder where a copy of a modified Common Component's Export
' file is kept, waiting for being moved to the Common Components folder - which is
' called "released". The class module maintains a 'Pending.dat' Private Profile
' file in the Pending folder for the release relevant properties:
' - ModInWbkName
' - ModInWbkFullName
' - ModExpFileFullName
' - ModAtDateTimeUTC
' - ModOnMachine
' - ExportFileName
' The class module provides the following methods:
' Register   Registers an Exportfile as pending release by copying it into the
'            Pending folder and updating the corresponding properties in the
'            Pending.dat Private Profile file.
' Hskpng     (Re-)Registers any used/hosted Common Component in the Serviced Workbook,
'            when identical with the already registered pending component.
' Remove     Removes a released component from the Pending dat Private Profile file,
'            when the last one registered, also removes the Pending folder.
' Release    Moves the pending release Export File from the Pending folder into the
'            Common Components folder by updating the corresponding Properties in the
'            CommComps.dat Private Profile file.
' Comps      Returns a Collection with the name of pending components (i.e. sections
'            in the Private Profile file Pending.dat)
' Conflict   Displays a dialog providing the following choices:
'            - Display former modifications
'            - Continue with former (pending release) modifications. I.e. the
'              pending release Export File is re-imported to the serviced Workbook in
'              order to enable continuation with ongoing modifications.
'            - Display this modifications
'            - Ignore the former modifications. I.e. this modifications become the
'              new pending release ones.
' -----------------------------------------------------------------------------------
Private Const MODIFIED_IN_WBK_NAME      As String = "Modified_InWorkbookName"
Private Const MODIFIED_IN_WBK_FULL_NAME As String = "Modified_InWorkbookFullName"
Private Const MODIFIED_EXP_FILE         As String = "Modified_ExportFileFullName"
Private Const MODIFIED_LAST_DATE_TIME   As String = "Modified_AtDateTime"
Private Const MODIFIED_ON_MACHINE       As String = "Modified_OnMachine"

Private CodePending As clsCode
Private CodePublic  As clsCode
Private PrivProf    As clsPrivProf
Private sFolder     As String
Private sCompName   As String

Private Sub Class_Initialize()

    Dim s As String
    
    '~~ Specify the pending release folder (exists only when in use)
    sFolder = wsConfig.FolderCommonComponentsPath & "\PendingReleases"
    
    '~~ Provide the Pending.dat Private Profile file's name
    Set PrivProf = New clsPrivProf
    PrivProf.FileName = wsConfig.FolderCommonComponentsPath & "\PendingReleases.dat"

End Sub

Public Sub DsplyDiffTheseModificationsVersusCurrentPublic(ByVal d_comp As clsComp)
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

Private Sub DsplyDiffTheseVersusFormerModifications(ByVal d_comp As clsComp)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
                    
    With d_comp
        .CodeCurrent.Source = .ExpFileFullName
        .CodePublic.Source = CommComps.LastModExpFileFullName(.CompName)
        .CodeCurrent.DsplyDiffs d_this_file_name:="ModificationsInThisWorkbook" _
                                 , d_this_file_title:="Current exported last modifications of the Common Component  " & mBasic.Spaced(sCompName) _
                                 , d_from_code:=CodePending _
                                 , d_from_file_name:="CurrentPendingReleaseModifications" _
                                 , d_from_file_title:="Modifications already made for Common Component  " & mBasic.Spaced(sCompName) & "   and pending release."
    End With
    
End Sub

Private Sub DsplyDiffFormerModificationsVersusCurrentPublic(ByVal d_comp As clsComp)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    With d_comp.CodePublic
        .DsplyDiffs d_this_file_name:="CurrentPublicCommonComponent" _
                  , d_this_file_title:="Current public code of the Common Component  " & mBasic.Spaced(sCompName) _
                  , d_from_code:=CodePending _
                  , d_from_file_name:="CurrentPendingReleaseModifications" _
                  , d_from_file_title:="Pending release modifications for Common Component  " & _
                                       mBasic.Spaced(sCompName) & "   already made in Workbook " & ModInWbkName
    End With
    
End Sub

Public Sub Management(ByVal p_comp As clsComp)
' ------------------------------------------------------------------------------
' - Remove obsolete 'Pending Releases' Common Components,
' - Establish/register 'Pending Releases' Common Components
' - Manage concurrent Common Component modification.
' ------------------------------------------------------------------------------
    Const PROC          As String = "Management"
    
    On Error GoTo eh
    Dim bModified       As Boolean
    Dim bNotModified    As Boolean
    Dim bPending        As Boolean
    Dim bNotPending     As Boolean
    Dim bDiffWbk        As Boolean
    Dim bEqulWbk        As Boolean
    Dim bDiffMachine    As Boolean
    Dim bEqulMachine    As Boolean

    With p_comp
        CompName = .CompName
        
        If CommComps.ExistsPublic(.CompName) Then
            .CodePublic.Source = CommComps.CurrentPublicExpFileFullName
            bModified = .CodeExported.DiffersFrom(.CodePublic, True)
        Else
            bModified = True
        End If
        bNotModified = Not bModified
        bPending = Exists(.CompName)
        bNotPending = Not bPending

        If bPending Then
            If CodePending Is Nothing Then Set CodePending = New clsCode
            CodePending.Source = ModExpFileFullName
            bDiffWbk = CommComps.LastModInWbkFullName <> vbNullString And CommComps.LastModInWbkFullName <> "unknown" And .WrkbkFullName <> ModInWbkFullName
            bEqulWbk = Not bDiffWbk
            bDiffMachine = ModOnMachine <> Environ("COMPUTERNAME")
            bEqulMachine = Not bDiffMachine
        End If

        ' Cases: -------------1-2-3-4-5-6--------------
        ' C1 Code modified *) y y y y y n
        ' C2 Pending          n y y y y y
        ' C3 Diff Wbk.        - n y n y n
        ' C4 Diff Machine.    - n n y y n
        ' ---------------------------------------
        ' A1 Reg. pend.       x
        ' A2 Reg. mod inf.      x
        ' A3 Manage conc.         x x x
        ' A4 Remove pending.            x
        ' ----------------------------------------
        ' *) The exported code differs from the public
        '    code in the Common-Components folder
        Select Case True
            Case bModified And bNotPending                                  ' Case 1        > A1
                Register p_comp
            Case bModified And bPending And bEqulWbk And bEqulMachine       ' Case 2        > A2
                If ModAtDateTimeUTC <> .LastModAtDateTimeUTC Then
                    ModAtDateTimeUTC = .LastModAtDateTimeUTC
                End If
                Debug.Print "Pending Release mod at updated: " & sCompName
            Case bModified And bPending And (bDiffWbk Or bDiffMachine)      ' Cases 3,4,5   > A3
                Conflict p_comp
            Case bNotModified And bPending And bEqulWbk And bEqulMachine    ' Case 6        > A4
                Remove .CompName
        End Select
    End With

xt: Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function Comps() As Collection
    Dim cll As Collection
    
    PrivProf.SectionNames names_cll:=cll
    Set Comps = cll

End Function

Public Sub Conflict(ByVal c_comp As clsComp)
' ------------------------------------------------------------------------------
' Displays a dialog providing the following choices:
' - Display these modifications (made in this Workbook)
' - Display the former modifications made elsewhere
' - Display theses versus the former modifications (see which modifications
'   had already been made versus those made in this Workbook)
' - Ignore these modifications. I.e. continue with former (pending release)
'   modifications.
' - Ignore the former modifications. I.e. these modifications become the new
'   pending release ones.
' ------------------------------------------------------------------------------
    Dim Msg As mMsg.udtMsg
    
    sCompName = c_comp.CompName
    
    With Msg
        With .Section(1)
            .Label.Text = "Concurrent modification"
            .Label.FontBold = True
            .Label.FontColor = rgbRed
            .Text.Text = _
            "Component:  """ & sCompName & """ has already been modified" & vbLf & _
            "in Workbook: " & ModInWbkName & vbLf & _
            "on computer: " & ModOnMachine & "." & vbLf & _
            "at date(utc): " & ModAtDateTimeUTC & vbLf & vbLf & _
            "When the modifications most recently made in" & vbLf & vbLf & _
            "this Workbook: " & ThisWorkbook.Name & vbLf & _
            "this computer: " & VBA.Environ$("computername") & vbLf & _
            "are confirmed, all the other changes made in the other Workbook are ignored. I.e. " & _
            "these modifications are given priority over the others already made. " & _
            "Comparing the modifications may support the decision on how to proceed."
        End With
        With .Section(2)
            .Label.Text = "Option Display this most recently made modifications:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Displays the difference between the current valid/released code versus these most resently made."
        End With
        With .Section(3)
            .Label.Text = "Option Display the already made modifications:"
            .Label.FontColor = rgbBlue
            .Text.Text = "Displays the difference between the current valid/released code versus those currently regigsterd waiting for rease."
        End With
        With .Section(4)
            .Label.Text = "Option Display the difference between the most recently made versus the already made modifications:"
            .Label.FontColor = rgbBlue
            .Text.Text = "This difference will enable to integrate (merge) the already made modifications with those most recently made in " & _
                         "order not to loose the others when finally the option ignore them is used."
        End With
        With .Section(5)
            .Text.Text = "Please note that one of the below two choices need to be made. While the first one implies the risk that the most " & _
                         "recently made modifications will get lost because not released, the second choice is final in the sense that the " & _
                         "already made modifications will get lost unless integrated/merged with these most curretly made ones."
        End With
        With .Section(6)
            .Label.Text = "Option Ignore the modifications made here and now:"
            .Label.FontColor = rgbBlue
            .Text.Text = "The changes already made in the other Workbook and/or the other machine remain those pending release." & vbLf & _
                         "This message will be repeated with each Workbook save event until either of the modifications has been released! " & _
                         "To be considered: Import the Export file with the other modifications and continue/finish the modifications in this VB-Project."
        End With
        With .Section(7)
            .Label.Text = "Option Ignore the other modifications already made before:"
            .Label.FontColor = rgbBlue
            .Text.Text = "The modifications made in this Workbook's VB-Project will become those pending release. I.e. the " & _
                         "other changes are ignored and will be replaced by thsi changes once released."
        End With
    End With
    Do
        Select Case mMsg.Dsply(dsply_title:="Severe danger of an already made code modification!" _
                             , dsply_msg:=Msg _
                             , dsply_Label_spec:="R130" _
                             , dsply_width_min:=30 _
                             , dsply_buttons:=mMsg.Buttons(BttnDsplyTheseModifications, BttnDsplyFormerModifications, vbLf, _
                                                           BttnDsplyDiffTheseVersusFormerModifications, vbLf, _
                                                           BttnIgnoreTheseModifications, BttnIgnoreFormerModifications))

            Case BttnDsplyTheseModifications:                   DsplyDiffTheseModificationsVersusCurrentPublic c_comp
            Case BttnDsplyFormerModifications:                  DsplyDiffFormerModificationsVersusCurrentPublic c_comp
            Case BttnDsplyDiffTheseVersusFormerModifications:   DsplyDiffTheseVersusFormerModifications c_comp
            Case BttnIgnoreFormerModifications:                 Register c_comp
                                                                Exit Do
            Case BttnIgnoreTheseModifications:                  Exit Do ' changes made in the other Workbook remain pending release
        End Select
    Loop
    
End Sub

Private Property Get BttnDsplyFormerModifications() As String
    BttnDsplyFormerModifications = "Display the modifications" & vbLf & "already made in Workbook" & vbLf & ModInWbkName & vbLf & "on computer " & vbLf & ModOnMachine
End Property

Private Property Get BttnDsplyTheseModifications() As String
    BttnDsplyTheseModifications = "Display these modifications" & vbLf & "made in this Workbook"
End Property

Private Property Get BttnDsplyDiffTheseVersusFormerModifications() As String
    BttnDsplyDiffTheseVersusFormerModifications = "Display modification difference" & vbLf & "this Workbook versus" & vbLf & "in " & ModInWbkName
End Property

Private Property Get BttnIgnoreFormerModifications() As String
    BttnIgnoreFormerModifications = "Ignore the other modifications" & vbLf & "made in Workbook" & vbLf & ModInWbkName & vbLf & "on computer " & vbLf & ModOnMachine
End Property

Private Property Get BttnIgnoreTheseModifications() As String
    BttnIgnoreTheseModifications = "Ignore the modifications" & vbLf & "made in this Workbook" & vbLf & Services.ServicedWbk.Name
End Property

Public Function Exists(ByVal e_comp As String)
' ------------------------------------------------------------------------------
' Returns TRUE when the component (e_comp) is pending release.
' ------------------------------------------------------------------------------
    With PrivProf
        Exists = .Exists(.FileName, e_comp)
    End With
End Function

Public Function Releases(Optional ByVal p_wbk As Workbook = Nothing) As Collection
' ------------------------------------------------------------------------------
' Returns a Collection of Common Components pending release.
' ------------------------------------------------------------------------------
    Const PROC = "PendingReleases"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim sComp       As String
    
    mBasic.BoP ErrSrc(PROC)
    Set Releases = New Collection
    For Each v In CommComps.Components
        sComp = v
        If Exists(sComp) _
        Then Releases.Add sComp
    Next v

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Register(ByVal r_comp As clsComp)
' -----------------------------------------------------------------------------------
' Registers an Export file as pending release by copying it into the Pending folder
' and updating the corresponding properties in the Pending.dat Private Profile file.
' Note: The PendingReleases folder exists only while there is at least one pending
'       release component.
' -----------------------------------------------------------------------------------
    
    With r_comp
        CompName = .CompName
        
        If Not FSo.FolderExists(sFolder) Then FSo.CreateFolder sFolder
        FSo.CopyFile .ExpFileFullName, sFolder & "\" & FSo.GetFileName(.ExpFileFullName), True
    
        ModExpFileFullName = .ExpFileFullName
        ModInWbkName = .Wrkbk.Name
        ModInWbkFullName = .Wrkbk.FullName
        ModAtDateTimeUTC = .LastModAtDateTimeUTC
    End With
    ModOnMachine = Environ("COMPUTERNAME")
    
End Sub

Public Sub ReleaseComp(ByVal r_comp As String)
' -----------------------------------------------------------------------------------
' 1. Move a component (r_comp) from the PendingRelease folder into the Common-Components
'    folder
' 2. Update the corresponding values in the CommComps.dat
' 3. Remove the corresponding values from the Pending.dat file - which by the way
'    removes the section
' -----------------------------------------------------------------------------------
    Const PROC = "ReleaseComp"
    
    On Error GoTo eh
    Dim sName   As String
    Dim sSource As String
    Dim sTarget As String
        
    mBasic.BoP ErrSrc(PROC)
    CompName = r_comp
    
    '~~ Move
    With FSo
        sName = .GetFileName(ModExpFileFullName)
        sSource = sFolder & "\" & sName
        sTarget = wsConfig.FolderCommonComponentsPath & "\" & sName
        If .FileExists(sTarget) Then .DeleteFile sTarget
        .MoveFile sSource, sTarget
    End With
    
    '~~ Update Common Component properties
    With CommComps
        .LastModExpFileFullNameOrigin = ModExpFileFullName
        .LastModInWbkName = ModInWbkName
        .LastModInWbkFullName = ModInWbkFullName
        .LastModAtDateTimeUTC = ModAtDateTimeUTC
    End With
    
    Remove r_comp
    
xt: Exit Sub
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
    
End Sub

Public Sub Remove(ByVal r_comp As String)
' -----------------------------------------------------------------------------------
' Removes the - component representing - section from the Private Profile file.
' When the removed section was the last one the Private Profile file will no longer
' exist after the removal and thus the Pending folder also has become obsolete and
' will be removed.
' -----------------------------------------------------------------------------------
    PrivProf.SectionRemove r_comp
    If Not FSo.FileExists(PrivProf.FileName) _
    Then FSo.DeleteFolder sFolder

End Sub
Private Sub Class_Terminate()
    Set PrivProf = Nothing
End Sub

Friend Property Let CompName(ByVal s As String)
    PrivProf.Section = s
    sCompName = s
End Property

Friend Property Get ModAtDateTimeUTC() As String:                       ModAtDateTimeUTC = Value(MODIFIED_LAST_DATE_TIME, sCompName):   End Property

Friend Property Let ModAtDateTimeUTC(ByVal p_mod_last As String):       Value(MODIFIED_LAST_DATE_TIME, sCompName) = p_mod_last:         End Property

Friend Property Get ModExpFileFullName() As String:                     ModExpFileFullName = Value(MODIFIED_EXP_FILE, sCompName):       End Property

Friend Property Let ModExpFileFullName(ByVal p_mod_exp_file As String): Value(MODIFIED_EXP_FILE, sCompName) = p_mod_exp_file:           End Property

Friend Property Get ModInWbkFullName() As String:                       ModInWbkFullName = Value(MODIFIED_IN_WBK_FULL_NAME, sCompName): End Property

Friend Property Let ModInWbkFullName(ByVal p_mod_by As String):         Value(MODIFIED_IN_WBK_FULL_NAME, sCompName) = p_mod_by:         End Property

Friend Property Get ModInWbkName() As String:                           ModInWbkName = Value(MODIFIED_IN_WBK_NAME, sCompName):          End Property

Friend Property Let ModInWbkName(ByVal p_mod_in_wbk_name As String):    Value(MODIFIED_IN_WBK_NAME, sCompName) = p_mod_in_wbk_name:     End Property

Friend Property Get ModOnMachine() As String:                           ModOnMachine = Value(MODIFIED_ON_MACHINE, sCompName):           End Property

Friend Property Let ModOnMachine(ByVal p_mod_at As String):             Value(MODIFIED_ON_MACHINE, sCompName) = p_mod_at:               End Property

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
    Value = PrivProf.Value(name_value:=v_value_name, name_section:=v_comp_name)

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
    PrivProf.Value(name_value:=v_value_name _
                 , name_section:=v_comp_name) = v_value
                                    
xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button
' - an "About:" section when the err_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' - the error message either by means of the Common VBA Message Service
'   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
'   means of the VBA.MsgBox in case not.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Consider extra information is provided with the error description
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsPendingRelease" & "." & sProc
End Function

