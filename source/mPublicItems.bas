Attribute VB_Name = "mPublicItems"
Option Explicit
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

'***App Window Constants***
Private Const WIN_NORMAL = 1         'Open Normal
Private Const WIN_MAX = 3            'Open Maximized
Private Const WIN_MIN = 2            'Open Minimized

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const vbResumeOk                As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
Private Const vbResume                  As Long = 6 ' return value (equates to vbYes)

Private dctCodeLines                    As Dictionary   ' All lines at first, finally those which refer to at least one public item
Private dctComponents                   As Dictionary
Private dctKindOfComponent              As Dictionary
Private dctKindOfItem                   As Dictionary
Private dctLinesReferringToPublicItems  As Dictionary
Private dctPublicItems                  As Dictionary   ' All Public ... and Friend ...
Private dctPublicXref                   As Dictionary   ' Public items with the Module.Procedures which refer to them
Private lLenPublicItems                 As Long
Private XrefExcluded()                  As Variant
Private XrefVBProject                   As VBProject
Private XrefWorkbook                    As Workbook
Private TotalCodeLines                  As Long

Private Property Get FileTemp(Optional ByVal tmp_path As String = vbNullString, _
                              Optional ByVal tmp_extension As String = ".tmp") As String
' ----------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file, when tmp_path
' is omitted in the CurDir path.
' ----------------------------------------------------------------------------
    Dim fso     As New FileSystemObject
    Dim sTemp   As String
    
    If VBA.Left$(tmp_extension, 1) <> "." Then tmp_extension = "." & tmp_extension
    sTemp = Replace(fso.GetTempName, ".tmp", tmp_extension)
    If tmp_path = vbNullString Then tmp_path = CurDir
    sTemp = VBA.Replace(tmp_path & "\" & sTemp, "\\", "\")
    FileTemp = sTemp
    Set fso = Nothing
    
End Property

Private Property Let FileText(Optional ByVal ft_file As Variant, _
                              Optional ByVal ft_append As Boolean = True, _
                              Optional ByRef ft_split As String, _
                                       ByVal ft_string As String)
' ----------------------------------------------------------------------------
' Writes the string (ft_string) into the file (ft_file) which might be a file
' object or a file's full name.
' Note: ft_split is not used but specified to comply with Property Get.
' ----------------------------------------------------------------------------
    Const PROC = "FileText-Let"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim ts  As TextStream
    Dim sFl As String
   
    ft_split = ft_split ' not used! just for coincidence with Get
    With fso
        If TypeName(ft_file) = "File" Then
            sFl = ft_file.Path
        Else
            '~~ ft_file is regarded a file's full name, created if not existing
            sFl = ft_file
            If Not .FileExists(sFl) Then .CreateTextFile sFl
        End If
        
        If ft_append _
        Then Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForAppending) _
        Else Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForWriting)
    End With
    
    ts.WriteLine ft_string

xt: ts.Close
    Set fso = Nothing
    Set ts = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get KindOfComponent(Optional ByVal comp_name As String) As String
    KindOfComponent = dctKindOfComponent(comp_name)
End Property

Private Property Let KindOfComponent(Optional ByVal comp_name As String, _
                                              ByVal comp_kind As String)
    dctKindOfComponent.Add comp_name, comp_kind
End Property

Private Property Get KindOfItem(Optional ByVal k_item As String) As String
    KindOfItem = dctKindOfItem(k_item)
End Property

Private Property Let KindOfItem(Optional ByVal k_item As String, _
                                  ByVal k_kind As String)
    If dctKindOfItem Is Nothing Then Set dctKindOfItem = New Dictionary
    If Not dctKindOfItem.Exists(k_item) Then dctKindOfItem.Add k_item, k_kind
End Property

Private Sub AddAscByKey(ByRef add_dct As Dictionary, _
                        ByVal add_key As Variant, _
                        ByVal add_item As Variant)
' ------------------------------------------------------------------------------------
' Adds to the Dictionary (add_dct) an item (add_item) in ascending order by the key
' (add_key). When the key is an object with no Name property an error is raisede.
'
' Note: This is a copy of the DctAdd procedure with fixed options which may be copied
'       into any VBProject's module in order to have it independant from this
'       Common Component.
'
' W. Rauschenberger, Berlin Jan 2022
' ------------------------------------------------------------------------------------
    Const PROC = "DAddAscByKey"
    
    On Error GoTo eh
    Dim bDone           As Boolean
    Dim dctTemp         As Dictionary
    Dim vItemExisting   As Variant
    Dim vKeyExisting    As Variant
    Dim vValueExisting  As Variant ' the entry's add_key/add_item value for the comparison with the vValueNew
    Dim vValueNew       As Variant ' the argument add_key's/add_item's value
    Dim bStayWithFirst  As Boolean
    Dim bOrderByItem    As Boolean
    Dim bOrderByKey     As Boolean
    Dim bSeqAscending   As Boolean
    Dim bCaseIgnored    As Boolean
    Dim bCaseSensitive  As Boolean
    Dim bEntrySequence  As Boolean
    
    If add_dct Is Nothing Then Set add_dct = New Dictionary
    
    '~~ Plausibility checks
    bOrderByItem = False
    bOrderByKey = True
    bSeqAscending = True
    bCaseIgnored = False
    bCaseSensitive = True
    bStayWithFirst = True
    bEntrySequence = False
    
    With add_dct
        '~~ When it is the very first add_item or the add_order option
        '~~ is entry sequence the add_item will just be added
        If .Count = 0 Or bEntrySequence Then
            .Add add_key, add_item
            GoTo xt
        End If
        
        '~~ When the add_order is by add_key and not stay with first entry added
        '~~ and the add_key already exists the add_item is updated
        If bOrderByKey And Not bStayWithFirst Then
            If .Exists(add_key) Then
                If VarType(add_item) = vbObject Then Set .Item(add_key) = add_item Else .Item(add_key) = add_item
                GoTo xt
            End If
        End If
    End With
        
    '~~ When the add_order argument is an object but does not have a name property raise an error
    If bOrderByKey Then
        If VarType(add_key) = vbObject Then
            On Error Resume Next
            add_key.Name = add_key.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(7), ErrSrc(PROC), "The add_order option is by add_key, the add_key is an object but does not have a name property!"
        End If
    ElseIf bOrderByItem Then
        If VarType(add_item) = vbObject Then
            On Error Resume Next
            add_item.Name = add_item.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(8), ErrSrc(PROC), "The add_order option is by add_item, the add_item is an object but does not have a name property!"
        End If
    End If
    
    vValueNew = AddAscByKeyValue(add_key)
    
    With add_dct
        '~~ Get the last entry's add_order value
        vValueExisting = AddAscByKeyValue(.Keys()(.Count - 1))
        
        '~~ When the add_order mode is ascending and the last entry's add_key or add_item
        '~~ is less than the add_order argument just add it and exit
        If bSeqAscending And vValueNew > vValueExisting Then
            .Add add_key, add_item
            GoTo xt
        End If
    End With
        
    '~~ Since the new add_key/add_item couldn't simply be added to the Dictionary it will
    '~~ be inserted before or after the add_key/add_item as specified.
    Set dctTemp = New Dictionary
    bDone = False
    
    For Each vKeyExisting In add_dct
        
        If VarType(add_dct.Item(vKeyExisting)) = vbObject _
        Then Set vItemExisting = add_dct.Item(vKeyExisting) _
        Else vItemExisting = add_dct.Item(vKeyExisting)
        
        With dctTemp
            If bDone Then
                '~~ All remaining items just transfer
                .Add vKeyExisting, vItemExisting
            Else
                vValueExisting = AddAscByKeyValue(vKeyExisting)
            
                If vValueExisting = vValueNew And bOrderByItem And bSeqAscending And Not .Exists(add_key) Then
                    If bStayWithFirst Then
                        .Add vKeyExisting, vItemExisting:   bDone = True ' not added
                    Else
                        '~~ The add_item already exists. When the add_key doesn't exist and bStayWithFirst is False the add_item is added
                        .Add vKeyExisting, vItemExisting:   .Add add_key, add_item:                     bDone = True
                    End If
                ElseIf bSeqAscending And vValueExisting > vValueNew Then
                    .Add add_key, add_item:                     .Add vKeyExisting, vItemExisting:   bDone = True
                Else
                    .Add vKeyExisting, vItemExisting ' transfer existing add_item, wait for the one which fits within sequence
                End If
            End If
        End With ' dctTemp
    Next vKeyExisting
    
    '~~ Return the temporary dictionary with the new add_item added and all exiting items in add_dct transfered to it
    Set add_dct = dctTemp
    Set dctTemp = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AddAscByKeyValue(ByVal add_key As Variant) As Variant
' ----------------------------------------------------------------------------
' When add_key is an object its name becomes the sort order value else the
' the value is returned as is.
' ----------------------------------------------------------------------------
    If VarType(add_key) = vbObject Then
        On Error Resume Next ' the object may not have a Name property
        AddAscByKeyValue = add_key.Name
        If Err.Number <> 0 Then Set AddAscByKeyValue = add_key
    Else
        AddAscByKeyValue = add_key
    End If
End Function

Private Sub AddPublicItem(ByVal pi_key As String, _
                          ByVal pi_item As String)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    lLenPublicItems = Max(lLenPublicItems, Len(pi_key))
    If Not dctPublicItems.Exists(pi_key) Then
        AddAscByKey dctPublicItems, pi_key, pi_item
    End If
End Sub

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error.
' - Returns a given positive 'Application Error' number (app_err_no) into a
'   negative by adding the system constant vbObjectError
' - Returns the original 'Application Error' number when called with a negative
'   error number.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub CodeLinesAndPublicItems()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC  As String = "CodeLinesAndPublicItems"
    
    On Error GoTo eh
    Dim vcmp    As VBComponent
    Dim v       As Variant
    Dim cmod    As CodeModule
    Dim sLine   As String
    Dim i       As Long
    Dim KoP     As vbext_ProcKind
    Dim sComp   As String
    Dim sLineId As String
    Dim sProc   As String
    Dim sItem   As String
    
    Set dctCodeLines = New Dictionary
    Set dctPublicItems = New Dictionary
    Set dctKindOfItem = New Dictionary
        
    BoP ErrSrc(PROC)
    For Each v In dctComponents
        sComp = v
'        Set cmod = dctComponents(v)
        Set vcmp = dctComponents(v)
        Set cmod = vcmp.CodeModule
        For i = 1 To cmod.CountOfLines
            sProc = cmod.ProcOfLine(i, KoP)
            sLine = Trim(cmod.Lines(i, 1))
'            If sLine = "Dim ErrMsgText  As TypeMsg" Then Stop
            Select Case True
                Case Len(sLine) <= 1                            ' Empty line
                Case LineBeginsWith(sLine, "Option Explicit")   ' Option line
                Case LineBeginsWith(sLine, "'")                 ' Comment line
                Case LineBeginsWith(sLine, "On Error")          ' On error line
                Case LineBeginsWith(sLine, "End ")              ' End
                Case LineBeginsWith(sLine, "Private ")          ' Private declaration line
                Case Else
                    sLineId = vcmp.Name & "." & sProc & "." & Format(i, "0000")
                    If Left(sLine, 7) = "Public " Or Left(sLine, 7) = "Friend " Then
                        UnstripComment sLine
                        Select Case True
                            Case IsPublicProc(sLine, sItem, sComp):                 AddPublicItem sComp & "." & sItem, sLineId & ": " & sLine
                            Case IsPublicEnum(sLine, sItem, sComp):                 AddPublicItem sComp & "." & sItem, sLineId & ": " & sLine
                            Case IsPublicType(sLine, sItem, sComp):                 AddPublicItem sComp & "." & sItem, sLineId & ": " & sLine
                            Case IsPublicConst(sLine, sItem, sComp):                AddPublicItem sComp & "." & sItem, sLineId & ": " & sLine
                            Case IsPublicDeclaration(sProc, sLine, sItem, sComp):   AddPublicItem sComp & "." & sItem, sLineId & ": " & sLine
                            Case Else
                                Debug.Print "Public item? " & sLine
                                Stop
                        End Select
                    Else
                        '~~ Collect all lines which do not specify or declare a Public or Friend item
                        UnstripComment sLine
                        dctCodeLines.Add sLineId, sLine
                    End If
            End Select
        Next i
    Next v

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
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
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
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
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal es_proc As String) As String
    ErrSrc = "mPublicItems" & "." & es_proc
End Function

Private Function IsExcluded(ByVal ie_comp As String) As Boolean
    Dim i   As Long
    Dim v   As Variant
    
    If VBA.IsArray(XrefExcluded) Then
        If UBound(XrefExcluded) >= 0 Then
            For Each v In XrefExcluded
                If v = ie_comp Then
                    IsExcluded = True
                    Exit Function
                End If
            Next v
        End If
    End If
    
End Function

Private Function IsPublicConst(ByVal is_line As String, _
                               ByRef is_item As String, _
                              ByVal is_comp As String) As Boolean
    If LineBeginsWith(is_line, "Public Const ", is_item) Then
        IsPublicConst = True
        KindOfItem(is_comp & "." & is_item) = "C"
    End If
    
End Function

Private Function IsPublicDeclaration(ByVal ipd_proc As String, _
                                     ByVal ipd_line As String, _
                                     ByRef ipd_item As String, _
                                     ByVal ipd_comp As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the code line (ipd_line) is a Public declaration. When the
' declaration specifies an instance of a Class Module the retunred item is
' <class-module>.<class-instance>
' ------------------------------------------------------------------------------
    Dim sInstance As String
    
    If ipd_proc = vbNullString And LineBeginsWith(ipd_line, "Public ", ipd_item) Then
        IsPublicDeclaration = True
        KindOfItem(ipd_comp & "." & ipd_item) = "V"
    End If
End Function

Private Function IsPublicEnum(ByVal is_line As String, _
                              ByRef is_item As String, _
                              ByVal is_comp As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the code line (ipd_line) is a Public Enum.
' ------------------------------------------------------------------------------
    If LineBeginsWith(is_line, "Public Enum ", is_item) Then
        IsPublicEnum = True
        KindOfItem(is_comp & "." & is_item) = "E"
    End If
End Function

Private Function IsPublicProc(ByVal is_line As String, _
                              ByRef is_item As String, _
                              ByVal is_comp As String) As Boolean
    Select Case True
        Case LineBeginsWith(is_line, "Public Function ", is_item)
            IsPublicProc = True
            KindOfItem(is_comp & "." & is_item) = "F"
        Case LineBeginsWith(is_line, "Public Sub ", is_item)
            IsPublicProc = True
            dctKindOfItem(is_comp & "." & is_item) = "S"
        Case LineBeginsWith(is_line, "Public Property Get ", is_item) _
          Or LineBeginsWith(is_line, "Public Property Let ", is_item) _
          Or LineBeginsWith(is_line, "Public Property Set ", is_item) _
          Or LineBeginsWith(is_line, "Friend Property Get ", is_item) _
          Or LineBeginsWith(is_line, "Friend Property Let ", is_item) _
          Or LineBeginsWith(is_line, "Friend Property Set ", is_item)
            IsPublicProc = True
            KindOfItem(is_comp & "." & is_item) = "P"
    End Select
    
End Function

Private Function IsPublicType(ByVal is_line As String, _
                              ByRef is_item As String, _
                              ByVal is_comp As String) As Boolean
    If LineBeginsWith(is_line, "Public Type ", is_item) Then
        IsPublicType = True
        dctKindOfItem(is_comp & "." & is_item) = "T"
    End If
End Function

Private Function IsSheet(ByVal is_comp_name As String, _
                         ByVal is_wb As Workbook) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Component's name (is_comp_name) is a Worksheet's
' CodeName in the Workbook (is_wb).
' ------------------------------------------------------------------------
    Dim ws As Worksheet
    
    For Each ws In is_wb.Worksheets
        If ws.CodeName = is_comp_name Then
            IsSheet = True
            Exit For
        End If
    Next ws

End Function

Private Function LineBeginsWith(ByVal lb_line As String, _
                                ByVal lb_string As String, _
                       Optional ByRef lb_item As String) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim lLen    As Long
    Dim sLine   As String
    Dim i       As Long
    Dim s       As String
    
    sLine = lb_line & " "
    If Left(sLine, Len(lb_string)) = lb_string Then
'        If lb_string = "Public Property Get " Then Stop
        LineBeginsWith = True
        If Left(lb_string, 7) = "Public " Then
            sLine = Replace(sLine, lb_string, vbNullString)
            If Len(sLine) > 1 Then
                For i = 1 To Len(sLine)
                    s = Mid(sLine, i, 1)
                    If s = "." Or s = "(" Or s = " " Then
                        lLen = i - 1
                        Exit For
                    End If
                Next i
                lb_item = Left(sLine, lLen)
            End If
        End If
    End If
        
End Function

Private Function Component(ByVal c_item As String) As String
    Component = Split(c_item, ".")(0)
End Function

Private Function CodeLineRefersToPublicItem(ByVal pi_line As String, _
                                            ByVal pi_item As String) As Boolean
' ------------------------------------------------------------------------------------
' Returns TRUE when the code line (pi_line) contains the 'Public Item' (pi_item).
' ------------------------------------------------------------------------------------
    Dim lStart      As Long   ' start pos
    Dim CharLeft    As String ' left of item char
    Dim CharRight   As String ' right of item char
    Dim sComp       As String
    Dim sItem       As String
    
    pi_line = " " & pi_line & " "
'    If InStr(pi_line, " TypeMsg ") <> 0 And pi_item = "TypeMsg" Then Stop
    lStart = InStr(pi_line, pi_item)
    If lStart <> 0 Then
        '~~ line possibly has a reference to component.item
        CharLeft = Mid(pi_line, lStart - 1, 1)
        If CharLeft = " " Or CharLeft = "(" Or CharLeft = "." Then
            CharRight = Mid(pi_line, lStart + Len(pi_item), 1)
            Select Case CharRight
                Case " ", "(", ")", ".", ","
                    CodeLineRefersToPublicItem = True
                    Exit Function
            End Select
        End If
    ElseIf UBound(Split(pi_item, ".")) >= 1 Then
        sItem = Split(pi_item, ".")(1)
        lStart = InStr(pi_line, sItem)
        If lStart <> 0 Then
            '~~ line possibly has a reference to the item
            CharLeft = Mid(pi_line, lStart - 1, 1)
            If CharLeft = " " Or CharLeft = "(" Or CharLeft = "." Then
                CharRight = Mid(pi_line, lStart + Len(sItem), 1)
                Select Case CharRight
                    Case " ", "(", ")", ".", ","
                        CodeLineRefersToPublicItem = True
                        Exit Function
                End Select
            End If
        End If
    End If
    
End Function

Private Sub LinesReferringToPublicItems()
' ------------------------------------------------------------------------------------
' Returns a Dictionary with Public Item as key and a Dictionaray with all lines
' rewfwerring to it as item whereby 'lines' are reduced to the <comonent>.procedure>.
' ------------------------------------------------------------------------------------
    Const PROC  As String = "LinesReferringToPublicItems"
    
    On Error GoTo eh
    Dim Items   As String
    Dim sComp   As String
    Dim sItem   As String
    Dim sLine   As String
    Dim sProc   As String
    Dim v1      As Variant
    Dim v2      As Variant
    Dim v3      As Variant
    Dim dct     As Dictionary
    Dim dctRefs As Dictionary
    Dim sRefId  As String
    Dim lItems  As Long
    Dim i       As Long
    
    BoP ErrSrc(PROC)
    Set dct = New Dictionary
    lItems = dctPublicItems.Count
    
    For Each v1 In dctPublicItems
        i = i + 1
        sComp = Split(v1, ".")(0)
        sItem = Split(v1, ".")(1)
        Set dctRefs = New Dictionary
'        Debug.Print Format(i, "000") & " of " & Format(lItems, "000") & ": " & v1
        DoEvents
        For Each v2 In dctCodeLines
            '~~ Loop through all code lines identifying those referring to the public item
            sLine = dctCodeLines(v2)
            If CodeLineRefersToPublicItem(sLine, v1) Then
                sRefId = Split(v2, ".")(0) & "." & Split(v2, ".")(1)
                If Not dctRefs.Exists(sRefId) And sRefId <> v1 Then
                    '~~ Keep record of the component.procedure referring to the public item
                    AddAscByKey dctRefs, sRefId, sLine
                End If
            End If
        Next v2
        dct.Add v1, dctRefs ' keep a record of the public item with the found lines referring to it
    Next v1
    
    Set dctLinesReferringToPublicItems = dct

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function Max(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the maximum value of all values provided (va).
' --------------------------------------------------------
    
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function OpenUrlEtc(ByVal oue_string As String, _
                            ByVal oue_show_how As Long) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples
' - Open a folder:          OpenUrlEtc("C:\TEMP\",WIN_NORMAL)
' - Call Email app:         OpenUrlEtc("mailto:dash10@hotmail.com",WIN_NORMAL)
' - Open URL:               OpenUrlEtc("http://home.att.net/~dashish", WIN_NORMAL)
' - Handle Unknown extensions (call Open With Dialog):
'                           OpenUrlEtc("C:\TEMP\TestThis",Win_Normal)
' - Start Access instance:  OpenUrlEtc("I:\mdbs\CodeNStuff.mdb", Win_NORMAL)
'
' Copyright:
' This code was originally written by Dev Ashish. It is not to be altered or
' distributed, except as part of an application. You are free to use it in any
' application, provided the copyright notice is left unchanged.
'
' Code Courtesy of: Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, oue_string, vbNullString, vbNullString, oue_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Error: Out of Memory/Resources. Couldn't Execute!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Error: File not found.  Couldn't Execute!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Error: Path not found. Couldn't Execute!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Error:  Bad File Format. Couldn't Execute!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & oue_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:         lRet = -1
    End Select
    
    OpenUrlEtc = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Public Sub Statistics(Optional ByVal xr_object As Variant = vbNullString, _
                      Optional ByVal xr_excluded As String = vbNullString)
' ------------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------------
    Const PROC  As String = "Statistics"
    
    On Error GoTo eh
    Dim sFile   As String
    Dim v1      As Variant
    Dim v2      As Variant
    Dim sItem   As String
    Dim dct     As Dictionary
    Dim sRefers As String
    
    BoP ErrSrc(PROC)
    References xr_object, xr_excluded ' provides dctPublicItems with all references in codelines
      
    sFile = FileTemp(tmp_extension:="txt")
    For Each v1 In dctPublicItems
        sItem = v1 & String(lLenPublicItems - Len(v1), ".") & " " & KindOfComponent(Component(v1)) & " " & KindOfItem(v1) & " "
        If dctLinesReferringToPublicItems.Exists(v1) Then
            Set dct = dctLinesReferringToPublicItems(v1)
            sRefers = sItem & String(3 - Len(CStr(dct.Count)), " ") & dct.Count
            FileText(sFile) = sRefers
        Else
            Debug.Print "No item found in dctLinesReferringToPublicItems for " & v1
        End If
    Next v1
    
xt: EoP ErrSrc(PROC)
#If ExecTrace = 1 Then
    OpenUrlEtc mTrc.LogFile, WIN_NORMAL
#End If
    OpenUrlEtc sFile, WIN_NORMAL
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Statistics_Test()
    Statistics Application.Workbooks("CompMan.xlsb"), "mBasic, mDct, mErH, mMsg, fMsg, mFile, mTrc, mWbk"
End Sub

Private Sub UnstripComment(ByRef uc_line As String)
' ------------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------------
    Dim l As Long
    l = InStr(uc_line, " '")
    If l <> 0 Then
        uc_line = Left(uc_line, l - 1)
    End If
End Sub

Private Sub XrefWbkAndVBP(ByVal xr_var As Variant)
' ------------------------------------------------------------------------------------
' Specifies XrefWorkbook and XrefVBProject
' ------------------------------------------------------------------------------------

    If TypeOf xr_var Is Workbook Then
        Set XrefWorkbook = xr_var
        Set XrefVBProject = XrefWorkbook.VBProject
    ElseIf TypeOf xr_var Is VBProject Then
        Set XrefVBProject = xr_var
    Else
        Set XrefWorkbook = ActiveWorkbook
        Set XrefVBProject = XrefWorkbook.VBProject
    End If
    
End Sub

Public Property Let Excluded(ByVal s As String)
    Dim v As Variant
    Dim i As Long
    
    If s <> vbNullString Then
        ReDim XrefExcluded(UBound(Split(s, ",")) + 1)
        XrefExcluded(0) = "mPublicItems" ' The component excludes itself
        For Each v In Split(s, ",")
            i = i + 1
            XrefExcluded(i) = Trim(v)
        Next v
    Else
        ReDim ExcludedComponents(0)
        ExcludedComponents(0) = "mPublicItems" ' The component excludes itself
    End If
    

End Property

Public Sub Xref(Optional ByVal xr_object As Variant = vbNullString, _
                Optional ByVal xr_excluded = vbNullString)
' ------------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------------
    Const PROC  As String = "Xref"
    
    On Error GoTo eh
    Dim sFile   As String
    Dim v1      As Variant
    Dim v2      As Variant
    Dim sItem   As String
    Dim dct     As Dictionary
    Dim sRefers As String
    
    BoP ErrSrc(PROC)
    References xr_object, xr_excluded ' provides dctPublicItems with all references in codelines
    
    sFile = FileTemp(tmp_extension:="txt")
    For Each v1 In dctPublicItems
        sItem = v1 & String(lLenPublicItems - Len(v1), ".") & ": "

        If dctLinesReferringToPublicItems.Exists(v1) Then
            Set dct = dctLinesReferringToPublicItems(v1)
            sRefers = sItem
            If dct.Count <> 0 Then
                For Each v2 In dct
                    sRefers = sRefers & v2 & vbLf & String(Len(sItem), " ")
                Next v2
            End If
            FileText(sFile) = sRefers
        Else
            Debug.Print "No item found in dctLinesReferringToPublicItems for " & v1
        End If
    Next v1
    
xt: EoP ErrSrc(PROC)
#If ExecTrace = 1 Then
    OpenUrlEtc mTrc.LogFile, WIN_NORMAL
#End If
    OpenUrlEtc sFile, WIN_NORMAL
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub References(Optional ByVal xr_object As Variant = vbNullString, _
                       Optional ByVal xr_excluded = vbNullString)
' ------------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------------
    Const PROC  As String = "References"
    
    On Error GoTo eh
    Dim dct     As Dictionary
    Dim sLine   As String
    Dim sItem   As String
    Dim v1      As Variant
    Dim v2      As Variant
    Dim sKey    As String
    Dim sRefers As String
    Dim i       As Long
    Dim vcmp    As VBComponent
    Dim cmod    As CodeModule
    Dim sFile   As String
        
    '~~ Preliminary preparations
    XrefWbkAndVBP xr_object
    Excluded = xr_excluded
    XrefComponents TotalCodeLines
    
#If ExecTrace = 1 Then
    mTrc.LogFile = Replace(XrefWorkbook.FullName, XrefWorkbook.Name, "Xref.trc")
    mTrc.LogTitle = "Cross-Reference for Public items in " & dctComponents.Count & " Components with a total of " & TotalCodeLines & " Code-Lines"
#End If

    BoP ErrSrc(PROC)
    CodeLinesAndPublicItems ' dctPublicItems, dctCodeLines, dctKindOrComponent, dctKindOfItem
    LinesReferringToPublicItems
          
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Xref_Test()
    Xref Application.Workbooks("CompMan.xlsb"), "mBasic, mDct, mErH, mMsg, fMsg, mFile, mTrc, mWbk"
End Sub

Private Sub XrefComponents(ByRef total_code_lines As Long)
' ------------------------------------------------------------------------------------
' Provides the Dictionary dctComponents with all not excluded component in the
' XrefVBProject and the total number of code lines in them.
' ------------------------------------------------------------------------------------
    Const PROC  As String = "XrefComponents"
    
    On Error GoTo eh
    Dim cmp     As VBComponent
    Dim sKind   As String
    
    BoP ErrSrc(PROC)
    Set dctComponents = New Dictionary
    Set dctKindOfComponent = New Dictionary
    
    For Each cmp In XrefVBProject.VBComponents
        With cmp
            Select Case .Type
                Case vbext_ct_ClassModule:  sKind = "Cls"
                Case vbext_ct_Document
                    If IsSheet(cmp.Name, XrefWorkbook) _
                    Then sKind = "Wsh" _
                    Else sKind = "Wbk"
                Case vbext_ct_MSForm:       sKind = "Frm"
                Case vbext_ct_StdModule:    sKind = "Mod"
            End Select
        
            If Not IsExcluded(.Name) Then
                AddAscByKey dctComponents, .Name, cmp
                KindOfComponent(cmp.Name) = sKind
                total_code_lines = total_code_lines + .CodeModule.CountOfLines
            Else
                Debug.Print "excluded: " & cmp.Name
            End If
        End With
    Next cmp

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
