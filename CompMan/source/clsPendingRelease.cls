VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPendingRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------------------------------------------------------
' Class Module clsPendingRelease: Fokus on maintainig Common Component modificatios
' =============================== pending release to the Common Componets folder
' (i.e. potentially unfinished modifications). The module mainains a 'Pending' folder
' within the Common Components folder and therein a 'Pending.dat' Private Profile
' file with the following values (each provided as a Get/Let Property):
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
' Comps      Returns a Dictionary of pending components (sections in the Pending.dat
'            Private Profile file)
' Conflicts  True when another modification (Workbook and/or machine) has a
'            registered pending release.
' Conflict   Displays a dialog with one of the following choices returned:
'            - Display former modifications
'            - Continue with former (pending release) modifications. I.e. the
'              pending release Export File is re-imported to the serviced Workbook in
'              order to enable continuation with ongoing modifications.
'            - Display this modifications
'            - Ignore the former modifications. I.e. this modifications become the
'              new pending release ones.
' -----------------------------------------------------------------------------------
Private Const EXPORT_FILE_NAME          As String = "ExportFileName"
Private Const MODIFIED_IN_WBK_NAME      As String = "Modified_InWorkbookName"
Private Const MODIFIED_IN_WBK_FULL_NAME As String = "Modified_InWorkbookFullName"
Private Const MODIFIED_EXP_FILE         As String = "Modified_ExportFileFullName"
Private Const MODIFIED_LAST_DATE_TIME   As String = "Modified_AtDateTime"
Private Const MODIFIED_ON_MACHINE       As String = "Modified_OnMachine"

Private PrivProf    As clsPrivProf
Private sFolder     As String
Private sCompName   As String

Private Sub Class_Initialize()

    Dim s As String
    
    '~~ Provide the pending release folder
    s = wsConfig.FolderCommonComponentsPath & "\PendingRelease"
    If Not FSo.FolderExists(s) Then FSo.CreateFolder s
    sFolder = s

    '~~ Provide the Pending.dat Private Profile file's name
    Set PrivProf = New clsPrivProf
    PrivProf.FileName = sFolder & "\Pending.dat"

End Sub

'Public Sub Management(ByVal p_comp As clsComp)
'' ------------------------------------------------------------------------------
'' Remove obsolete 'Pending Releases' and establish/register 'Pending Releases'
'' or manage a concurrent modification.
'' ------------------------------------------------------------------------------
'    Dim bModified       As Boolean
'    Dim bNotModified    As Boolean
'    Dim bPending        As Boolean
'    Dim bNotPending     As Boolean
'    Dim bDiffWbk        As Boolean
'    Dim bEqulWbk        As Boolean
'    Dim bDiffMachine    As Boolean
'    Dim bEqulMachine    As Boolean
'
'    With p_comp
'        CompName = .CompName
'        If ExpFileFullName <> vbNullString Then
'            .CodePublic.Source = CommComps.CurrentPublicExpFileFullName
'            bModified = .CodeExported.DiffersFrom(p_comp.CodePublic, True)
'        Else
'            bModified = True
'        End If
'            bNotModified = Not bModified
'        bPending = CommComps.PendingReleaseRegistered
'        bNotPending = Not bPending
'
'        If bPending Then
'            CodePending.Source = PendingReleaseModExpFileFullName
'            bDiffWbk = LastModInWbkFullName <> vbNullString And LastModInWbkFullName <> "unknown" And .WrkbkFullName <> PendingReleaseModInWbkFullName
'            bEqulWbk = Not bDiffWbk
'            bDiffMachine = PendingReleaseModOnMachine <> vbNullString _
'                       And PendingReleaseModOnMachine <> Environ("COMPUTERNAME")
'            bEqulMachine = Not bDiffMachine
'        End If
'    End With
'
'    ' Cases: -------------1-2-3-4-5-6--------------
'    ' C1 Code modified *) y y y y y n
'    ' C2 Pending          n y y y y y
'    ' C3 Diff Wbk.        - n y n y n
'    ' C4 Diff Machine.    - n n y y n
'    ' ---------------------------------------
'    ' A1 Reg. pend.       x
'    ' A2 Reg. mod inf.      x
'    ' A3 Manage conc.         x x x
'    ' A4 Remove pending.            x
'    ' ----------------------------------------
'    ' *) The exported code differs from the public
'    '    code in the Common-Components folder
'    Select Case True
'        Case bModified And bNotPending                                  ' Case 1        > A1
'            PendingReleaseRegister = p_comp
'        Case bModified And bPending And bEqulWbk And bEqulMachine       ' Case 2        > A2
'            If PendingReleaseModAtDateTimeUTC <> p_comp.LastModAtDateTimeUTC Then
'                PendingReleaseModAtDateTimeUTC = p_comp.LastModAtDateTimeUTC
'            End If
'            Debug.Print "Pending Release mod at updated: " & sCompName
'        Case bModified And bPending And (bDiffWbk Or bDiffMachine)      ' Cases 3,4,5   > A3
'            ConcurrentModification p_comp
'        Case bNotModified And bPending And bEqulWbk And bEqulMachine    ' Case 6        > A4
'            PendingReleaseRemove
'    End Select
'
'End Sub

Public Function Comps() As Dictionary:  Set Comps = PrivProf.SectionNames():    End Function

Public Sub ReleaseComp(ByVal r_comp_name As String)
' -----------------------------------------------------------------------------------
' 1. Move a component (r_comp) from the PendingRelease folder into the Common-Components
'    folder
' 2. Update the corresponding values in the CommComps.dat
' 3. Remove the corresponding values from the Pending.dat file - which by the way
'    removes the section
' -----------------------------------------------------------------------------------
    Const PROC = "ReleaseComp"
    
    On Error GoTo eh
    
    CompName = r_comp_name
    MoveToCommonComponentsFolder PendingReleaseModExpFileFullName
    With CommComps
        .LastModExpFileFullNameOrigin = PendingReleaseModExpFileFullName
        .LastModInWbkName = PendingReleaseModInWbkName
        .LastModInWbkFullName = PendingReleaseModInWbkFullName
        .LastModAtDateTimeUTC = PendingReleaseModAtDateTimeUTC
    End With
    
    PrivProf.ValueRemove name_value:=MODIFIED_ON_MACHINE & vbCrLf & _
                                     MODIFIED_IN_WBK_NAME & vbCrLf & _
                                     MODIFIED_IN_WBK_FULL_NAME & vbCrLf & _
                                     MODIFIED_LAST_DATE_TIME & vbCrLf & _
                                     MODIFIED_EXP_FILE _
                       , name_section:=r_comp_name

xt: Exit Sub
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
    
End Sub

Friend Property Let ExportFileName(Optional ByVal l_comp_name As String = vbNullString, _
                                            ByVal l_ext As String)
    PrivProf.Value(EXPORT_FILE_NAME, sCompName) = l_ext
End Property
                                                 
Friend Property Get ExportFileName(Optional ByVal l_comp_name As String = vbNullString) As String
    ExportFileName = PrivProf.Value(EXPORT_FILE_NAME, sCompName)
End Property

Private Sub MoveToCommonComponentsFolder(ByVal m_comp As String)
' ------------------------------------------------------------------------------
' Save the Export-File (s_file) of a pending release to the Common-Components
' folder.
' Note: This is exclusively used by the mRelease.ReleaseComp procedure which
'       subsequently updates the corresponding properties.
' ------------------------------------------------------------------------------
    Const PROC  As String = "MoveToCommonComponentsFolder"
        
    On Error GoTo eh
    Dim fl          As File
    Dim sFileSource As String
    Dim sFileTarget As String
    
    mBasic.BoP ErrSrc(PROC)
    
'    With FSo
'        If TypeName(s_file) = "File" _
'        Then Set fl = s_file _
'        Else Set fl = .GetFile(s_file)
'        sFileName = .GetFileName(fl)
'        sFileFullName = fl.Path
'
'        .MoveFile sFileFullName, wsConfig.FolderCommonComponentsPath & "\" & sFileName
'        If .GetExtensionName(fl.Name) = "frm" Then
'            '~~ When the Export file has a .frm extension the .frx file needs to be copied too
'            .CopyFile Replace(sFileFullName, ".frm", ".frx"), wsConfig.FolderCommonComponentsPath & "\" & .GetBaseName(fl) & ".frx"
'        End If
'    End With
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Class_Terminate()
    Set PrivProf = Nothing
End Sub

Friend Property Let CompName(ByVal s As String)
    PrivProf.Section = s
    sCompName = s
End Property

Friend Property Get Folder() As String
    Dim s As String
    s = wsConfig.FolderCommonComponentsPath & "\PendingRelease"
    If Not FSo.FolderExists(s) Then FSo.CreateFolder s
    sFolder = s
    
End Property

Friend Property Get PendingReleaseModAtDateTimeUTC(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModAtDateTimeUTC = Value(MODIFIED_LAST_DATE_TIME, p_comp_name)
End Property

Friend Property Let PendingReleaseModAtDateTimeUTC(Optional ByVal p_comp_name As String = vbNullString, _
                                                         ByVal p_mod_last As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(MODIFIED_LAST_DATE_TIME, p_comp_name) = p_mod_last
End Property

Friend Property Get PendingReleaseModExpFileFullName(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModExpFileFullName = Value(MODIFIED_EXP_FILE, p_comp_name)
End Property

Friend Property Let PendingReleaseModExpFileFullName(Optional ByVal p_comp_name As String = vbNullString, _
                                                              ByVal p_mod_exp_file As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(MODIFIED_EXP_FILE, p_comp_name) = p_mod_exp_file
End Property

Friend Property Get PendingReleaseModInWbkFullName(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModInWbkFullName = Value(MODIFIED_IN_WBK_FULL_NAME, p_comp_name)
End Property

Friend Property Let PendingReleaseModInWbkFullName(Optional ByVal p_comp_name As String = vbNullString, _
                                                            ByVal p_mod_by As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(MODIFIED_IN_WBK_FULL_NAME, p_comp_name) = p_mod_by
End Property

Friend Property Get PendingReleaseModInWbkName(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModInWbkName = Value(MODIFIED_IN_WBK_NAME, p_comp_name)
End Property

Friend Property Let PendingReleaseModInWbkName(Optional ByVal p_comp_name As String = vbNullString, _
                                                        ByVal p_mod_in_wbk_name As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(MODIFIED_IN_WBK_NAME, p_comp_name) = p_mod_in_wbk_name
End Property

Friend Property Get PendingReleaseModOnMachine(Optional ByVal p_comp_name As String = vbNullString) As String
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    PendingReleaseModOnMachine = Value(MODIFIED_ON_MACHINE, p_comp_name)
End Property

Friend Property Let PendingReleaseModOnMachine(Optional ByVal p_comp_name As String = vbNullString, _
                                                 ByVal p_mod_at As String)
    If p_comp_name = vbNullString Then p_comp_name = sCompName
    Value(MODIFIED_ON_MACHINE, p_comp_name) = p_mod_at
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

