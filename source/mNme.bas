Attribute VB_Name = "mNme"
Option Explicit
' ------------------------------------------------------------------------
' Standard Module mNme: Services about Name objects.
'
' Public services:
' - Corresponding   Returns a Dictionary with all Name objects in a
'                   provided Workbook which correspond with a provided
'                   Name object whereby 'corresponding' means with either
'                   the same name and/or the same referred range.
' - Create          Returns a Name object with provided properties and the
'                   provided scope.
' - Exists          Returns TRUE and the resulting name object when a
'                   Range Name exists in a provided Workbook disregarding
'                   any difference in the RefersTo argument.
' - HasChanged      Returns TRUE when a provided Name object refers to
'                   a range in a Workbook which has exactly one but
'                   another name.
' - IsNotUnique     Returns TRUE when a Name refers to a range in a
'                   Workbook which has already one but another name.
' - IsInUse         Returns TRUE when a provided Name's Name is found in
'                   any of a Workbook's VBComponent code lines.
' - MereName        Returns the mere Name of a Name object, i.e. with a
'                   scope sheet prefix unstripped.
' - Remove          Removes a provided Name or all Names which refer to a
'                   provided Worksheet
' - Scope           Returns the scoped Workbook or Worksheet object for a
'                   provided Name object.
' - ScopeName       Returns a provided Name object's scope as string
'                   "scope: Workbook" or "scope: Worksheet <name>"
' - UnifiedId       Returns a unified id for a provided Name object in the
'                   form: <name> | <refersto> | scope: Workbook
'                      or <name> | <refersto> | scope: Worksheet <name>
'                   either "Workbook scope" or <Worksheet.Name> scope.

'
' W. Rauschenberger, Berlin, Dec 2022
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
    sName = mNme.MereName(stb_nme)                 ' Save the mere name of the name
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

Public Sub Remove(ByVal r_var As Variant)
' ----------------------------------------------------------------------------
' Removes a named Name object when the argument is a string or all Name
' which refer to a range of the Worksheet when the argument is a Worksheet.
' Note: A worksheet as argument may serve the need to remove all corresponding
'       Names before the Worksheet gets deleted - in order to prevent invalid
'       Names reulting from the deletion of the Worksheet.
' ----------------------------------------------------------------------------
    Dim wsh As Worksheet
    Dim wbk As Workbook
    Dim nme As Name
    
    Select Case VarType(r_var)
        Case vbString
            '~~ The provided argument is considered a Name objects Name
        Case vbObject
            '~~ The provided argument should be a Worksheet
            If TypeOf r_var Is Worksheet Then
                Set wsh = r_var
                Set wbk = wsh.Parent
                For Each nme In wbk.Names
                    If InStr(nme.RefersTo, "=" & wsh.Name & "!") <> 0 Then
                        nme.Delete
                    End If
                Next nme
            End If
    End Select
End Sub

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
' Common 'Begin of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.        | Installed component |
'         |-------------------------|---------------------|
'         | ErHComp = 1             | mErH                |
'         | XcTrc_mTrc = 1          | mTrc                |
'         | XcTrc_clsTrc = 1        | clsTrc              |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
    Dim s As String
    If Not IsMissing(b_arguments) Then s = Join(b_arguments, ";")

#If ErHComp = 1 Then
    '~~ The error handling will also hand over to the Common VBA Execution Trace
    '~~ provided one is installed (mTrc/clsTrc) and activated.
    mErH.BoP b_proc, s
#ElseIf XcTrc_clsTrc = 1 Then
    '~~ mErH is not installed but the mTrc is
    Trc.BoP b_proc, s
#ElseIf XcTrc_mTrc = 1 Then
    '~~ mErH neither mTrc is installed but clsTrc is
    mTrc.BoP b_proc, s
#End If

End Sub

Private Sub EoP(ByVal e_proc As String, Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.        | Installed component |
'         |-------------------------|---------------------|
'         | ErHComp = 1             | mErH                |
'         | XcTrc_mTrc = 1          | mTrc                |
'         | XcTrc_clsTrc = 1        | clsTrc              |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ The error handling will also hand over to the Common VBA Execution Trace
    '~~ provided one is installed (mTrc/clsTrc) and activated.
    mErH.EoP e_proc
#ElseIf XcTrc_clsTrc = 1 Then
    Trc.EoP e_proc, e_inf
#ElseIf XcTrc_mTrc = 1 Then
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
    If err_source = vbNullString Then err_source = Err.source
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
    ErrSrc = "mNme." & s
End Function

Public Function Corresponding(ByVal cn_nme As Name, _
                              ByVal cn_wbk As Workbook, _
                     Optional ByRef cn_corresponding As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with Name object in the provided Workbook (cn_wbk)
' which correspond with the provided Name object (cn_nme). Name objects with
' an invalid Name or invalid RefersTo property are ignored. The additionally
' returned Dictionary (cn_corresponding) allows a usage like the following:
'
' Dim dct As Dictionary
' If CorrespondingNames(nme, wbk, dct).Count <> 0 Then
'     For each v in dct
'         ....
'     Next v
'
' An error is raised when the parent Workbook of the provided Name object
' (cn_nme) and the provided Workbook (cn_wbk) is the same Workbook.
'
' Note: The error handling is to be provided by the caller.
' ----------------------------------------------------------------------------
    Const PROC = "Corresponding"
    
    Dim nme             As Name
    Dim dct             As New Dictionary
    Dim sId             As String
    Dim wbk             As Workbook
    Dim bCorrespName    As Boolean
    Dim bCorrespRange   As Boolean
    Dim nmeCorresp      As Name
    
    For Each nme In cn_wbk.Names
        sId = UnifiedId(nme)
        Set nmeCorresp = nme
        bCorrespName = mNme.MereName(nme) = mNme.MereName(cn_nme) And InStr(nme.Name, "#") = 0 And InStr(cn_nme.Name, "#") = 0
        bCorrespRange = nme.RefersTo = cn_nme.RefersTo And InStr(nme.RefersTo, "#") = 0 And InStr(cn_nme.RefersTo, "#") = 0
        If bCorrespName Or bCorrespRange Then
            If Not dct.Exists(sId) Then dct.Add sId, nmeCorresp
        End If
    Next nme
    Debug.Print dct.Count & ". " & dct.Keys()(dct.Count - 1) & ": " & dct.Items()(dct.Count - 1).Name
    
    Set Corresponding = dct
    Set cn_corresponding = dct
    Set dct = Nothing
    
End Function

Public Function ScopeName(ByVal sn_nme As Name) As String
    
    If TypeOf sn_nme.Parent Is Worksheet _
    Then ScopeName = "scope: Worksheet" _
    Else ScopeName = "scope: Workbook"
    ScopeName = ScopeName & " '" & sn_nme.Parent.Name & "'"
    
End Function

Public Function UnifiedId(ByVal ui_nme As Name, _
                 Optional ByVal ui_delim As String = " | ") As String
' ----------------------------------------------------------------------------
' Returns a unified id for a Name object (ui_nme) in the form:
' <name>-<refersto> when the scope is Workbook or
' <name>-<refersto>-<scopedsheetname> when the scope is Worksheet.
'
' Note: A Worksheet scoped Name object (ui_nme) may not conform with the sheet
'       the Name the ReferTo property specifies.
' ----------------------------------------------------------------------------
    Dim s As String
    UnifiedId = MereName(ui_nme) & ui_delim & ui_nme.RefersTo
    s = ScopeName(ui_nme)
    If s <> vbNullString Then UnifiedId = UnifiedId & ui_delim & s
End Function

Public Function Exists(ByVal ex_nme As Name, _
                       ByVal ex_wbk As Workbook, _
              Optional ByRef ex_nme_result As Name) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when a Name object exists in the provided Workbook (ex_wbk)
' which corresponds with the provided Name object whereby 'corresponding'
' means: exactly one Name object with either the same Name or the same
' referred range.
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

Public Function HasChangedName(ByVal hc_nme As Name, _
                               ByVal hc_wbk As Workbook, _
                      Optional ByRef hc_nme_result As Name) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Name objects (hc_nme_source) in the Workbook
' (hc_wbk_source) corresponds with exactly one Name object in the Workbook
' (hc_wbk) but this one has a different Name.
' ------------------------------------------------------------------------
    Const PROC = "HasChangedName"
    
    On Error GoTo eh
    Dim nme As Name
    Dim dct As New Dictionary
    
    If mNme.Corresponding(hc_nme, hc_wbk, dct).Count = 1 Then
        Set nme = dct.Items()(0)
        HasChangedName = nme.Name <> hc_nme
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), "The Name object has either no or more than one corresponding " & _
                                           "Name objects in the provided Workbook '" & hc_wbk.Name & "'!"
    End If
    
xt: Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Create(ByVal c_name As String, _
                       ByVal c_rng As Range, _
                       ByVal c_scope As Variant, _
              Optional ByRef c_nme As Name) As Name
' ------------------------------------------------------------------------
' Returns a Name object, optionally (c_nme), with the provided properties
' and the provided scope (c_scope).
' ------------------------------------------------------------------------
    Const PROC = "Create"
    
    Dim nme     As Name
    Dim Scope As Variant
    Dim wbk As Workbook
    Dim wsh As Worksheet
         
    BoP ErrSrc(PROC)
    If TypeOf c_scope Is Workbook Then
        Set wbk = c_scope
        Set nme = wbk.Names.Add(c_name, c_rng)
    ElseIf TypeOf c_scope Is Worksheet Then
        Set wsh = c_scope
        Set nme = wsh.Names.Add(c_name, c_rng)
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), "The provided scope (c_scope) is neither a Workbook nor a Worksheet!"
    End If
        
    Set c_nme = nme
    Set Create = nme
    Set nme = Nothing
    EoP ErrSrc(PROC)
    
End Function

Public Function HasChangedReferredRange(ByVal hc_nme As Name, _
                                        ByVal hc_wbk As Workbook, _
                               Optional ByRef hc_nme_result As Name) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Name (hc_nme) refers to another range than the -
' unique single corresponding name in the Workbook (hc_wbk).
' Precondition: Equal named Worksheets. I.e. the referred sheet in the
'               Name (hc_nme) and the corresponding sheet in Workbook
'               (hc_wbk) have equal names.
' ------------------------------------------------------------------------
    Const PROC = "HasChangedReferredRange"
    
    On Error GoTo eh
    Dim nme As Name
    
    For Each nme In hc_wbk.Names
        If nme.Name = hc_nme.Name Then
            Set hc_nme_result = nme
            HasChangedReferredRange = nme.RefersTo <> hc_nme.RefersTo
            Exit For
        End If
    Next nme
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
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

Public Function MereName(ByVal no_nme As Name) As String
' ------------------------------------------------------------------------------
' Returns a Name object's mere Name, i.e. one without a Sheetname prefix
' ------------------------------------------------------------------------------
    Dim v As Variant
    v = Split(no_nme.Name, "!")
    MereName = v(UBound(v))
End Function

Public Sub ChangeProperties(ByRef p_target_nme As Name, _
                   Optional ByVal p_source_nme As Name = Nothing, _
                   Optional ByVal p_final_name As String = vbNullString, _
                   Optional ByVal p_final_rng As Range = Nothing, _
                   Optional ByVal p_final_scope As Variant = Nothing, _
                   Optional ByRef p_name_changed As Boolean, _
                   Optional ByRef p_range_changed As Boolean, _
                   Optional ByRef p_scope_changed As Boolean)
' ------------------------------------------------------------------------------
' The service returns():
' - the Name object (p_target_nme) with the properties provided either by a
'   Name object (p_source_nme) - i.e. synchronized - or alternatively with
'   properties changed by explicitely provided (p_final_nmae, p_final_range,
'   p_final_scope). A provided Scope (p_final_scope) is either a Workbook or a
'   Worksheet object of the target Workbook.
' - TRUE for those properties which had been changed (p_name_changed,
'   p_range_changed, p_scope_changed)
'
' - Usage rules:
' - The following error conditions need to be handled by the caller):
'   - The provided target Name object and the provided source Name object
'     (p_source_nme) are from the same Workbook
'   - The referred Range sheet name is not a sheet name in the target Name
'     object's Workbook
'   - The final scope is Worksheet and a sheet of that name does not exist in
'   target Name object's Workbook
'
' Note 1: When a source Name object is provided any explicitely provided
'         arguments are ignored.
' Note 2: The service has been provided to synchronize a target Name object
'         with a source Name object but may as well be used th change any
'         property including the scope.
'
' W. Rauschenberger Berlin, Dec 2022
' ------------------------------------------------------------------------------
    Const PROC = "ChangeProperties"
    
    On Error GoTo eh
    Dim FinalName           As String
    Dim FinalNme            As Name
    Dim FinalRefersTo       As String
    Dim FinalRngAddress     As String
    Dim FinalRngSheet       As String
    Dim FinalScope          As Variant ' either a Workbook or a Worksheet object
    Dim FinalScopeSheet     As String
    Dim FinalWbk            As Workbook
    Dim OldName             As String
    Dim OldRefersTo         As String
    Dim OldScopeName        As String
    Dim SourceScopeSheet    As String
    Dim SourceWbk           As Workbook
    Dim wsh                 As Worksheet
    
    BoP ErrSrc(PROC)
    OldName = mNme.MereName(p_target_nme)
    OldRefersTo = p_target_nme.RefersTo
    OldScopeName = ScopeName(p_target_nme)
    
    '~~ Final Name property
    If Not p_source_nme Is Nothing Then
        FinalName = p_source_nme.Name
    Else
        If p_final_name <> vbNullString _
        Then FinalName = p_final_name _
        Else FinalName = mNme.MereName(p_target_nme)
    End If
    
    '~~ Final Sheet Name and Range Address
    If Not p_source_nme Is Nothing Then
        FinalRngAddress = p_source_nme.RefersToRange.Address
        FinalRngSheet = p_source_nme.RefersToRange.Parent.Name
    Else
        If Not p_final_rng Is Nothing Then
            FinalRngAddress = p_final_rng.Address
            FinalRngSheet = p_final_rng.Parent.Name
        Else
            FinalRngAddress = p_target_nme.RefersToRange.Address
            FinalRngSheet = p_target_nme.RefersToRange.Parent.Name
        End If
    End If
        
    '~~ Final Scope
    If Not p_source_nme Is Nothing Then
        Set FinalScope = p_source_nme.Parent
    Else
        If Not p_final_scope Is Nothing _
        Then Set FinalScope = p_final_scope _
        Else Set FinalScope = p_target_nme.Parent
    End If
    If TypeOf FinalScope Is Worksheet Then FinalScopeSheet = FinalScope.Name
   
    '~~ Final Workbook
    If TypeOf p_target_nme.Parent Is Workbook _
    Then Set FinalWbk = p_target_nme.Parent _
    Else Set FinalWbk = p_target_nme.Parent.Parent
       
    '~~ Error Conditions
    If Not p_source_nme Is Nothing Then
        If TypeOf p_source_nme.Parent Is Workbook _
        Then Set SourceWbk = p_source_nme.Parent _
        Else Set SourceWbk = p_source_nme.Parent.Parent
        If SourceWbk.FullName = FinalWbk.FullName _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided source Name's and the provided target Name's Workbook must not be the same!"
    End If
    If Not SheetExists(FinalWbk, FinalRngSheet) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The final referred range sheet's name (" & FinalRngSheet & ") does not exist in the target Name object's Worbook (" & FinalWbk.Name & ")!"
    If TypeOf FinalScope Is Worksheet Then
        If Not SheetExists(FinalWbk, FinalScopeSheet) _
        Then Err.Raise AppErr(2), ErrSrc(PROC), "The target Name's Workbook does not have a Worksheet named '" & FinalScopeSheet & "'!"
    End If
    
    FinalRefersTo = "=" & FinalRngSheet & "!" & FinalRngAddress
    
    '~~ Provide the final Name object depending on the FinalScope
    Select Case True
        Case (TypeOf FinalScope Is Workbook And TypeOf p_target_nme.Parent Is Workbook) _
          Or (TypeOf FinalScope Is Worksheet And TypeOf p_target_nme.Parent Is Worksheet And FinalScope.Name = p_target_nme.Parent.Name)
            '~~ Scope not changed
            Set FinalNme = p_target_nme
        Case Else
            p_target_nme.Delete
            '~~ When the scope has to be changed a new target Name object has to be created which finally replaces the provided
            If TypeOf FinalScope Is Workbook Then
                Set FinalNme = FinalWbk.Names.Add(FinalName, FinalRefersTo)
            ElseIf TypeOf FinalScope Is Worksheet Then
                Set wsh = FinalWbk.Worksheets(FinalScopeSheet)
                Set FinalNme = wsh.Names.Add(FinalName, FinalRefersTo)
            End If
    End Select
    
    FinalNme.Name = FinalName
    FinalNme.RefersTo = FinalRefersTo
    Set p_target_nme = FinalNme
    
    '~~ Compile changed id
    p_name_changed = mNme.MereName(FinalNme) <> OldName
    p_range_changed = FinalNme.RefersTo <> OldRefersTo
    p_scope_changed = ScopeName(FinalNme) <> OldScopeName
    Set FinalNme = Nothing
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function SheetExists(ByVal sx_wbk As Workbook, _
                             ByVal sx_sheet_name As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when a sheet named (sx_sheet_name) exists in Workbook (sx_wbk).
' ------------------------------------------------------------------------------
    Dim wsh As Worksheet
    
    On Error Resume Next
    Set wsh = sx_wbk.Worksheets(sx_sheet_name)
    SheetExists = Err.Number = 0
    
End Function

Private Function FoundInFormulas(ByVal fif_str As String, _
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
    Dim Rng As Range
    
    BoP PROC
    For Each wsh In fif_wbk.Worksheets
        If Not fif_wsh Is Nothing Then
            If Not wsh Is fif_wsh Then GoTo ws
        End If
        
        On Error Resume Next
        Set Rng = wsh.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Err.Number <> 0 Then GoTo ws
        For Each cel In Rng
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
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FoundInCodeLines(ByVal ficl_str As String, _
                                  ByVal ficl_wbk As Workbook, _
                         Optional ByRef ficl_cll As Collection) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the range name (ficl_nme) is used by any of the
' VBComponent's code lines in the workbook (ficl_wbk) or in any formula
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
' VBComponent's code lines in the workbook (iu_wbk) or in any formula
' in any Worksheet's cell.
' ------------------------------------------------------------------------
    Const PROC = "IsInUse"
    
    On Error GoTo eh
    Dim s   As String
    
    s = mNme.MereName(iu_nme)
    IsInUse = FoundInFormulas(s, iu_wbk, , iu_cll)
    If Not IsInUse Then IsInUse = FoundInCodeLines(s, iu_wbk, iu_cll)
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
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

