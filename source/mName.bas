Attribute VB_Name = "mName"
Option Explicit
' ------------------------------------------------------------------------
' Standard Module mName
'
' Public services:
' - Exists          Returns TRUE and the resulting name object when a Range
' Name (ex_nme) exists in the Workbook (ex_wbk) - disregarding any
' difference in the RefersTo argument.
' - HasChanged      Returns TRUE when a provided Name object refers to
'                   a range in a Workbook which has exactly one but
'                   another name.
' - IsNotUnique     Returns TRUE when a Name refers to a range in a
'                   Workbook which has already one but another name.
' - IsInUse         Returns TRUE when a provided Name's Name is found in
'                   any of a Workbooks VBComponent's code lines.
'
' W. Rauschenberger, Berlin, Sept 2022
' ------------------------------------------------------------------------
Public Const SCOPE_DELIM = "|"

Public Enum enNameScope
    enWorkbook
    enWorksheet
End Enum

Public Sub WshScopeToWbkScope(ByVal cs_wbk As Workbook, _
                     Optional ByVal cs_nme As Variant = vbNullString)
' ----------------------------------------------------------------------------
' Change the scope from Worksheet scope to Workbook (cs_wbk) scope.
' - When a Name object is provided (cs_nme) the scope is changed for the Name,
'   whereby provided Name (cs_nme) is used only for the mere name and the
'   range the Name referres to. I.e. the provided Workbook (cs_wbk) is the new
'   scope regardless of the scope of the provided Name as long as it is a
'   Worksheet scope
' - When a filter (cs_filter) is provided, the scope change is performed for
'   all names in either in the Workbook or any of its Worksheets of which the
'   Name property meets the filter string (supports all "Like" placeholders).
'   One should keep in mind the fact that a Name may be prefixed by the
'   Worksheet's name it is in scope. In order to assure all concerned Names
'   are covered the filter needs to be "*" only.
' ----------------------------------------------------------------------------
    Const PROC = "WshScopeToWbkScope"
    
    On Error GoTo eh
    Dim nme As Name
    Dim wsh As Worksheet
    
    Select Case True
        Case TypeName(cs_nme) = "Name"
            SheetToBookScope cs_nme, cs_wbk
        Case TypeName(cs_nme) = "String"
            For Each nme In cs_wbk.Names
                If cs_nme = vbNullString Then
                    SheetToBookScope cs_nme, cs_wbk
                ElseIf nme.Name Like cs_nme Then
                    SheetToBookScope cs_nme, cs_wbk
                End If
            Next nme
            For Each wsh In cs_wbk.Worksheets
                For Each nme In wsh.Names
                    If cs_nme = vbNullString Then
                        SheetToBookScope cs_nme, cs_wbk
                    ElseIf nme.Name Like cs_nme Then
                        SheetToBookScope cs_nme, cs_wbk
                    End If
                Next nme
            Next wsh
    End Select

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub SheetToBookScope(ByVal stb_nme As Name, _
                             ByVal stb_wbk As Workbook)
' ----------------------------------------------------------------------------
' Provided the Name (stb_nme) is in the scope of a Worksheet - regardless of
' which Worksheet - the Name is removed and recreates with Workbook (stb_wbk)
' scope.
' ----------------------------------------------------------------------------
    Dim sName   As String
    Dim sRef    As String
    
    If Not TypeOf stb_nme.Parent Is Worksheet Then Exit Sub
    sName = mName.Mere(stb_nme)                 ' Save the mere name of the name
    sRef = stb_nme.RefersTo                         ' Save the referred range
    stb_nme.Delete                                  ' Delete the name
    
    stb_wbk.Names.Add Name:=sName, RefersTo:=sRef   ' Create a new Name with Workbook scope

End Sub

Public Function ScopeIsWorkbook(ByVal si_nme As Name) As Boolean
    ScopeIsWorkbook = TypeOf si_nme.Parent Is Workbook
End Function

Public Function ScopeIsWorkSheet(ByVal si_nme As Name, _
                        Optional ByRef si_wsh_name As String) As Boolean
' ----------------------------------------------------------------------------
' When the Name (si_nme) is in the scope of a Worksheet the function returns
' TRUE and the scoped Worksheet's name (si_wsh_name).
' ----------------------------------------------------------------------------
    If TypeOf si_nme.Parent Is Worksheet Then
        ScopeIsWorkSheet = True
        si_wsh_name = si_nme.Parent.Name
    End If
End Function

Public Function Scope(ByVal scp_nme As Name, _
             Optional ByRef scp_wbk As Workbook, _
             Optional ByRef scp_wsh As Worksheet) As enNameScope
' ----------------------------------------------------------------------------
' Returns the name of the scope object of the provided Name (sc_nme) as String.
' A Workbook's Name when the scope is Workbook
' or a <workbook-name>|<worksheet-name> when the scope is Worksheet. by the
' way the Workbook (sc_nme_result_wbk) and if applicable the Worksheet
' (sc_nme_result_wbk) is returned as objects.
' ----------------------------------------------------------------------------
    Select Case True
        Case TypeOf scp_nme.Parent Is Workbook
            Set scp_wbk = scp_nme.Parent
            Scope = enWorkbook
        Case TypeOf scp_nme.Parent Is Worksheet
            Set scp_wsh = scp_nme.Parent
            Set scp_wbk = scp_wsh.Parent
            Scope = enWorksheet
    End Select

End Function
                          
Private Sub AddAscByKey(ByRef add_dct As Dictionary, _
                           ByVal add_key As Variant, _
                           ByVal add_item As Variant)
' ----------------------------------------------------------------------------
' Adds to the Dictionary (add_dct) an item (add_item) in ascending order by
' the key (add_key). When the key is an object with no Name property an error
' is raised.
'
' Note: This is a copy of the DctAdd procedure with fixed options which may be
'       copied into any VBProject's module in order to have it independant
'       from this Common Component.
'
' W. Rauschenberger, Berlin Jan 2022
' ----------------------------------------------------------------------------
    Const PROC = "DctAdd"
    Dim bDone           As Boolean
    Dim dctTemp         As Dictionary
    Dim vItem           As Variant
    Dim vItemExisting   As Variant
    Dim vKeyExisting    As Variant
    Dim vValueExisting  As Variant ' the entry's add_key/add_item value for the comparison with the vValueNew
    Dim vValueNew       As Variant ' the argument add_key's/add_item's value
    Dim vValueTarget    As Variant ' the add before/after add_key/add_item's value
    Dim bStayWithFirst  As Boolean
    Dim bOrderByItem    As Boolean
    Dim bOrderByKey     As Boolean
    Dim bSeqAscending   As Boolean
    Dim bCaseIgnored    As Boolean
    Dim bCaseSensitive  As Boolean
    Dim bEntrySequence  As Boolean
    
    On Error GoTo eh
    
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
' Universal error message display service. See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
'
' Basic service:
' - Displays a debugging option button when the Conditional Compile Argument
'   'Debugging = 1'
' - Displays an optional additional "About the error:" section when a string is
'   concatenated with the error message by two vertical bars (||)
' - Displays the error message by means of VBA.MsgBox when neither of the
'   following is installed
'
' Extendend service when other Common Components are installed and indicated via
' Conditional Compile Arguments:
' - Invokes mErH.ErrMsg when the Conditional Compile Argument ErHComp = 1
' - Invokes mMsg.ErrMsg when the Conditional Compile Argument MsgComp = 1 (and
'   the mErH module is not installed / MsgComp not set)
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to turn
'          them into negative and in the error message back into a positive
'          number.
' - ErrSrc To provide an unambiguous procedure name by prefixing is with the
'          module name.
'
' See: https://github.com/warbe-maker/Common-VBA-Error-Services
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf MsgComp = 1 Then
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
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

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

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mName." & s
End Function

Public Function Exists(ByVal ex_nme As Name, _
                       ByVal ex_wbk As Workbook, _
              Optional ByRef ex_nme_result As Name) As Boolean
' ------------------------------------------------------------------------
' When the Range Name (ex_nme) exists in the Workbook (ex_wbk) -
' disregarding any difference in the RefersTo argument - the function
' returns TRUE and the resulting name object (ex_nme_result).
' ------------------------------------------------------------------------
    Dim nme As Name
    
    For Each nme In ex_wbk.Names
        If nme.Name = ex_nme.Name Then
            Exists = True
            Exit Function
        End If
    Next nme
        
End Function

Public Function IsValidUserRangeName(ByVal iv_nme As Name) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided Name (iv_nme) is a valid user range name.
' ----------------------------------------------------------------------------
    IsValidUserRangeName = iv_nme.RefersTo <> vbNullString And iv_nme.Name <> vbNullString
    If IsValidUserRangeName _
    Then IsValidUserRangeName = Not iv_nme.Name Like "_xlfn.*" _
                            And InStr(iv_nme.RefersTo, "=") <> 0 _
                            And InStr(iv_nme.RefersTo, "!") <> 0 _
                            And InStr(iv_nme.RefersTo, "$") <> 0 _
                            And iv_nme.RefersTo <> "=#NAME?"
End Function

Public Function HasChanged(ByVal hc_nme_source As Name, _
                           ByVal hc_wbk_target As Workbook, _
                           ByVal hc_wbk_source As Workbook, _
                  Optional ByRef hc_nme_source_result As Name) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Name (hc_nme_source) in the Workbook
' (hc_wbk_source) refers to a Range which in in the Workbook
' (hc_wbk_target) has exactly one but another name.
' ------------------------------------------------------------------------
    Const PROC = "HasChanged"
    
    On Error GoTo eh
    Dim nmeTarget   As Name
    Dim dct         As New Dictionary
    
    Set dct = mRng.HasNames(hc_nme_source.RefersTo)
    If dct.Count = 1 Then
        HasChanged = Not dct.Exists(hc_nme_source.Name)
    End If
    
xt: Set dct = Nothing
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

                        
Public Function HasMoved(ByVal hc_nme As Name, _
                         ByVal hc_wbk As Workbook, _
                Optional ByRef hc_nme_result As Name) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Name (hc_nme_source) refers to another range
' than the same name in the Workbook (hc_wbk).
' Precondition: Equal named Worksheets. I.e. the referred sheet in the
'               Name (hc_nme) and the corresponding sheet in Workbook
'               (hc_wbk) have equal names.
' ------------------------------------------------------------------------
    Const PROC = "HasMoved"
    
    On Error GoTo eh
    Dim nme As Name
    
    For Each nme In hc_wbk.Names
        If nme.Name = hc_nme.Name Then
            Set hc_nme_result = nme
            HasMoved = nme.RefersTo <> hc_nme.RefersTo
            Exit For
        End If
    Next nme
    
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsNotUnique(ByVal ia_nme As Name, _
                            ByVal ia_wbk As Workbook) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when at least one Name  in Workbook (in_wbk) refers to the
' same range as the Name (in_nme) but none is identical with the Name
' (ia_nme).
' ------------------------------------------------------------------------
    Dim dct As Dictionary:  Set dct = mRng.HasNames(ia_nme.RefersTo)
    
    IsNotUnique = dct.Count > 0 And Not dct.Exists(ia_nme.Name)

End Function

Public Function Mere(ByVal no_nme As Name) As String
' ------------------------------------------------------------------------------
' Returns a Name objects mere Name, i.e. one without a Sheetname prefix
' ------------------------------------------------------------------------------
    Dim v As Variant
    v = Split(no_nme.Name, "!")
    Mere = v(UBound(v))
    
End Function

Public Function FoundInFormulas(ByVal fif_str As String, _
                                ByVal fif_wbk As Workbook, _
                       Optional ByVal fif_wsh As Worksheet = Nothing, _
                       Optional ByRef fif_cll As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the provided string (fit_str) is found in a formula. When a
' Worksheet (fif_wsh) is provided, only in this else in all Worksheets in the
' provided Workbook (fif_wbk). When a Collection (fif_cll) is provided all cells
' with a formula containing the string (fif_str) are returned, else the function
' ends with TRUE with the first found.
' ------------------------------------------------------------------------------
    Const PROC = "FoundInFormulas"
    
    On Error GoTo eh
    Dim cel As Range
    Dim cll As New Collection
    Dim wsh As Worksheet
    Dim rng As Range
    
    BoP PROC
    For Each wsh In fif_wbk.Worksheets
        If Not fif_wsh Is Nothing Then
            If Not wsh Is fif_wsh Then GoTo ws
        End If
        
        On Error Resume Next
        Set rng = wsh.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Err.Number <> 0 Then GoTo ws
        For Each cel In rng
            If InStr(1, cel.Formula, fif_str) > 0 Then
                FoundInFormulas = True
                If IsMissing(fif_cll) Then
                    GoTo xt
                Else
                    cll.Add cel
                End If
            End If
        Next cel
        
ws: Next wsh

xt: Set fif_cll = cll
    Set cll = Nothing
    EoP PROC
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FoundInCodeLines(ByVal ficl_str As String, _
                                 ByVal ficl_wbk As Workbook, _
                        Optional ByRef ficl_cll As Collection) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the range name (ficl_nme) is used by any of the
' VB-Component's code lines in the workbook (ficl_wbk) or in any formula
' in any Worksheet's cell.
' ------------------------------------------------------------------------
    Const PROC = "FoundInCodeLines"
    Const C_APOST = "FoundInCodeLines"

    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim i       As Long
    Dim sLine   As String
    Dim vbcm    As CodeModule
    
    For Each vbc In ficl_wbk.VBProject.VBComponents
        Set vbcm = vbc.CodeModule
        With vbcm
            For i = 1 To .CountOfLines
                sLine = .Lines(i, 1)
                If Not IsOutCommented(sLine, ficl_str) Then
                    If InStr(C_APOST & sLine & C_APOST, ficl_str) <> 0 Then
                        FoundInCodeLines = True
                        If ficl_cll Is Nothing Then
                            Exit Function
                        Else
                            ficl_cll.Add vbc.Name & ": " & i & ": " & sLine
                        End If
                    End If
                End If
            Next i
        End With
    Next vbc
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsInUse(ByVal iu_nme As Name, _
                        ByVal iu_wbk As Workbook, _
               Optional ByRef iu_cll As Collection) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the range name (iu_nme) is used by any of the
' VB-Component's code lines in the workbook (iu_wbk) or in any formula
' in any Worksheet's cell.
' ------------------------------------------------------------------------
    Const PROC = "IsInUse"
    
    On Error GoTo eh
    IsInUse = FoundInFormulas(mName.Mere(iu_nme), iu_wbk, , iu_cll)
    If Not IsInUse Then IsInUse = FoundInCodeLines(mName.Mere(iu_nme), iu_wbk, iu_cll)
    
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function IsOutCommented(ByVal ioc_line As String, _
                                ByVal ioc_item As String) As Boolean
    Dim i As Long
    For i = InStr(ioc_line, ioc_item) To 1 Step -1
        If VBA.Mid(ioc_line, i, 1) = "'" Then
            IsOutCommented = True
            Exit For
        End If
    Next i
    
End Function

