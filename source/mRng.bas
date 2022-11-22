Attribute VB_Name = "mRng"
Option Explicit
Option Compare Text
' ------------------------------------------------------------------------------
' Standard  Module mRng Provides services for ranges.
'
' Methods:
' - Exists          Returns TRUE when the range exists
' - FoundInFormulas Returns TRUE when a provided string is found in any formula.
'                   When a Worksheet is provided, only in this else in all
'                   Worksheets in the provided Workbook. When a Collection is
'                   provided all cells with a formula containing the provided
'                   string are returned, else the function ends with TRUE with
'                   the first found.
' - FromArray       not yet implemented
' - HasNames        Returns a Dictionary with all Names referring to a certain
'                   range.
' - HasUrl
' - Raised          Mimiks a cell as a not clicked button
' - Sunken          Mimiks a cell as a clicked button
' - Url Get         Returns the hyperling of a provided range/cell
' - UrlLet          Inserts a Hyperlink into a provided range/cell
'
' Uses:     No other modules
'
' Requires: Reference to "Microsoft Scripting Runtine"
'
' W. Rauschenberger, Berlin Oct 2022
' ------------------------------------------------------------------------------

Public Property Get Url(ByVal u_rng As Range) As String
' ------------------------------------------------------------------------------
' Returns the "Full" Hyperlink Address for the range (u_rng) even when the
' system just provides a relative address. When the range (u_rng) does not
' contain a hyperlink a vbNullString is returned.
' ------------------------------------------------------------------------------
    
    Dim fso         As New FileSystemObject
    Dim wbk         As Workbook
    Dim sUrl        As String
    Dim sSplit      As String
    Dim aUrlHost    As Variant
    Dim aUrlPath    As Variant
    Dim i           As Long
    Dim sRoot       As String
    Dim sPath       As String
    
    Set wbk = u_rng.Parent.Parent
    sPath = wbk.Path
    fso.GetParentFolderName (sPath)
    
    aUrlHost = Split(wbk.FullName, "\")
    
    If u_rng.Hyperlinks.Count = 1 Then
        sUrl = u_rng.Hyperlinks(1).Address
        If InStr(sUrl, "/") <> 0 Then sSplit = "/" Else sSplit = "\"
        aUrlPath = Split(sUrl, sSplit)
        If UBound(aUrlPath) = 0 Then
            sUrl = aUrlHost(LBound(aUrlHost))
            For i = LBound(aUrlHost) + 1 To UBound(aUrlHost) - 1
                sUrl = sUrl & sSplit & aUrlHost(i)
            Next i
            Url = sUrl & sSplit & aUrlPath(0)
        ElseIf aUrlPath(0) = ".." Then
            '~~ Handle ..
            sUrl = vbNullString
            For i = UBound(aUrlPath) To LBound(aUrlPath) Step -1
                If aUrlPath(i) = ".." Then
                    sUrl = sUrl & sSplit & fso.GetParentFolderName(sPath)
                    sPath = fso.GetParentFolderName(sPath)
                End If
            Next i
            Debug.Print sUrl
            
            For i = LBound(aUrlPath) + 1 To UBound(aUrlPath)
                If aUrlPath(i) <> ".." Then
                    sUrl = sUrl & sSplit & aUrlPath(i)
                End If
            Next i
            Url = sUrl
        Else
            If u_rng.Hyperlinks(1).SubAddress <> vbNullString Then
                Url = sUrl & "#" & u_rng.Hyperlinks(1).SubAddress
            Else
                Url = sUrl
            End If
        End If
    End If
    Url = Replace(Url, sSplit, vbNullString, 1, 1)
    
End Property

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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRng" & "." & sProc
End Function

Public Sub SetBackColor(ByVal sbc_rng As Range, _
                        ByVal sbc_rbg As XlRgbColor)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    With sbc_rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .Color = sbc_rbg
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    
End Sub
                        
Public Function HasName(ByVal hn_rng As Range) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the range (hn_rng) has a Name.
' ------------------------------------------------------------------------------
    Dim nme As name
    Dim s   As String
    
    On Error Resume Next
    Set nme = Nothing ' only including in case you are planning to use this approach in a loop.
    Set nme = hn_rng.name
    On Error GoTo xt
    If Not nme Is Nothing Then
        s = nme.name
        HasName = True
    End If
    
xt: Exit Function

End Function

Public Sub IndicateByBackColor()
    Dim rgbInput            As XlRgbColor
    Dim rgbLockedNamed      As XlRgbColor
    Dim rgbLockedUnNamed    As XlRgbColor
    
    rgbInput = RGB(242, 242, 242)
    rgbLockedNamed = RGB(255, 255, 204)
    
    rgbLockedUnNamed = rgbGray
    
End Sub

Public Function Exists(ByVal e_wbk As Variant, _
                       ByVal e_wsh As Variant, _
                       ByVal e_rng As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the provided Range (e_rng) - which may be range object or a
' range's name - exists in the provided Worksheet (e_wsh) - which may be a
' sheet's name or a Worksheet object - of the provided Workbook (e_wbk) - which
' may be a Workbbok's Name or Full-Name or a Workbook object.
' ------------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim sTest       As String
    Dim wsh         As Worksheet
    Dim wbk         As Workbook
    Dim Rng         As Range
    Dim wsResult    As Worksheet

    BoP PROC
    
    Exists = False
    Set wbk = mCompMan.WbkGetOpen(e_wbk)   ' raises an error when not open
    
    If TypeName(e_rng) = "Nothing" _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The Range (parameter e_rng) for the Range's existence check is ""Nothing""!"
    
    If TypeOf e_rng Is Range Then
        Set Rng = e_rng
        On Error Resume Next
        sTest = Rng.Address
        Exists = Err.Number = 0
    ElseIf VarType(e_rng) = vbString Then
        If Not SheetExists(wbk, e_wsh, wsResult) _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The Worksheet (parameter vWs) for the Range's existence check does not exist in Workbook (e_wbk)!"
        Set wsh = wsResult
        On Error Resume Next
        sTest = wsh.Range(e_rng).Address
        Exists = Err.Number = 0
    End If
            
xt: EoP PROC
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

'Public Sub FromArray(ByVal r As Range, _
'                     ByVal vArr As Variant)
'' -----------------------------------------
'' Returns the range (r) filled with the
'' content of array (vArr).
'' -----------------------------------------
'
'End Sub

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
    
    Dim cel As Range
    Dim cll As New Collection
    Dim wsh As Worksheet
    Dim Rng As Range
    
    BoP PROC
    For Each wsh In fif_wbk.Worksheets
        If Not fif_wsh Is Nothing Then
            If Not wsh Is fif_wsh Then GoTo nw
        End If
        
        On Error Resume Next
        Set Rng = wsh.UsedRange.SpecialCells(xlCellTypeFormulas)
        If Err.Number <> 0 Then GoTo nw
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
        
nw: Next wsh

xt: Set fif_cll = cll
    Set cll = Nothing
    EoP PROC
    Exit Function

End Function

Public Function HasNames(ByVal hn_rng As Range) As Dictionary
' ------------------------------------------------------------------------
' Returns a Dictionary with all Names for the range (hn_rng).
' ------------------------------------------------------------------------
    Dim nme As name
    Dim dct As New Dictionary
    Dim wbk As Workbook
    
    Set wbk = hn_rng.Parent.Parent
    For Each nme In wbk.Names
        If nme.RefersTo = hn_rng.name Then
            dct.Add nme.name, nme.RefersTo
        End If
    Next nme
    
    Set HasNames = dct
    Set dct = Nothing

End Function

Public Function HasUrl(ByVal r As Range) As Boolean
    HasUrl = r.Hyperlinks.Count > 0
End Function

Public Sub Raised(ByVal r As Range)
    With r
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThick
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThick
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.399945066682943
            .PatternTintAndShade = 0
        End With
    End With
End Sub

Private Function SheetExists(ByVal vWb As Variant, _
                             ByVal vWs As Variant, _
                    Optional ByRef wsResult As Worksheet) As Boolean
' ------------------------------------------------------------
' Returns TRUE when the Worksheet (vWs) - which may be a
' Worksheet object or a Worksheet's name - exists in the
' Workbook (vWb).
' ------------------------------------------------------------
    Const PROC  As String = "SheetExists"    ' This procedure's name for the error handling and execution tracking
    
    On Error GoTo eh
    Dim wsTest  As Worksheet
    Dim wb      As Workbook
    Dim ws      As Worksheet

    SheetExists = False
    
    If TypeName(vWb) = "Nothing" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter vWb) is ""Nothing""!"
    
    If Not WbkIsOpen(vWb, wb) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "The Workbook (parameter vWb) is not open!"

    If TypeName(vWs) = "Nothing" _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The Worksheet (parameter vWs) for the Worksheet's existence check is ""Nothing""!"
    
    If Not TypeOf vWs Is Worksheet And VarType(vWs) <> vbString _
    Then Err.Raise AppErr(4), ErrSrc(PROC), "The Worksheet (parameter vWs) for the Worksheet's existence check is neither a Worksheet object nor a Worksheet's name or modulename!"
    
    If TypeOf vWs Is Worksheet Then
        Set ws = vWs
        For Each wsTest In wb.Worksheets
            If wsTest Is ws Then
                SheetExists = True
                Set wsResult = wsTest
                GoTo xt
            End If
        Next wsTest
        GoTo xt
    ElseIf VarType(vWs) = vbString Then
        For Each wsTest In wb.Worksheets
            If wsTest.name = vWs Then
                SheetExists = True
                Set wsResult = wsTest
                GoTo xt
            End If
        Next wsTest
    End If
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Sunken(ByVal r As Range)
    With r
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThick
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.499984740745262
            .Weight = xlThick
        End With
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = 0.399945066682943
            .PatternTintAndShade = 0
        End With
    End With
End Sub

'Public Sub ToArray(ByVal r As Range, _
'                   ByRef vArr As Variant)
'' -----------------------------------------
'' Returns the array (vArr) filled with the
'' content of the range (r).
'' -----------------------------------------
'
'End Sub

Public Sub UrlLet(ByVal u_wsh As Worksheet, _
                  ByVal u_rng As Range, _
                  ByVal u_address As String, _
         Optional ByVal u_text_to_display As String = vbNullString)
' --------------------------------------------------------------
'
' --------------------------------------------------------------
    u_rng.Hyperlinks.Delete
    u_wsh.Hyperlinks.Add Anchor:=u_rng, Address:=u_address, TextToDisplay:=u_text_to_display
End Sub

Public Function WbkGetOpen(ByVal go_wb As Variant, _
                  Optional ByVal go_read_only As Boolean = False) As Workbook
' ----------------------------------------------------------------------------
' Returns an open Workbook object or raises an error. If go_wb is a full path-file name, the file exists but
' is not open it is opened.
' Note: A ReadOnly mode has to be set by the caller.
' ----------------------------------------------------------------------------
    Const PROC = "GetOpen"
    
    On Error GoTo eh
    Dim sWbName     As String
    Dim sWbFullName As String
    Dim wbOpen      As Workbook
    Dim fso         As New FileSystemObject
        
    Set wbOpen = Nothing
       
    If mWbk.IsWbObject(go_wb) Then
        sWbName = go_wb.name
        sWbFullName = go_wb.FullName
    ElseIf mWbk.IsFullName(go_wb) Then
        sWbName = fso.GetFileName(go_wb)
        sWbFullName = go_wb
    ElseIf mWbk.IsName(go_wb) Then
        sWbName = go_wb
        sWbFullName = vbNullString
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter go_wb) is neither a Workbook object nor a string (name or fullname)!" & vbLf & _
                                           "(TypeName of argument = '" & TypeName(go_wb) & "'!)"
    End If

    If mWbk.IsWbObject(go_wb) Then
        Set WbkGetOpen = go_wb
        GoTo xt
    End If
    
    With mWbk.Opened
        If .Exists(sWbName) Then
            Set WbkGetOpen = .Item(sWbName)
            If sWbFullName = vbNullString Then GoTo xt
            Set wbOpen = .Item(sWbName)
            If wbOpen.FullName = sWbFullName Then GoTo xt
            '~~ A Workbook with the Name is open but it has a different FullName
            If fso.FileExists(sWbFullName) Then
                '~~ When the Workbook file still exists at the provided location the one which is open is the wromg one
                Err.Raise AppErr(2), ErrSrc(PROC), "A Workbook named '" & sWbName & "' is open but its location differs." & vbLf & vbLf & _
                                                   "'" & wbOpen.FullName & "'" & vbLf & vbLf & _
                                                   "instead of the requested" & vbLf & vbLf & _
                                                   "'" & sWbFullName & "'"
            End If
            GoTo xt
        Else
            '~~ A Workbook with the provided name is yet not open
            If sWbFullName = vbNullString _
            Then Err.Raise AppErr(3), ErrSrc(PROC), "A Workbook named '" & sWbName & "' is not open - and cannot be opened because this requires the full file name!"
            If Not fso.FileExists(sWbFullName) _
            Then Err.Raise AppErr(4), ErrSrc(PROC), "A Workbook named '" & sWbFullName & "' does not exist!"
            Set WbkGetOpen = Workbooks.Open(sWbFullName, , go_read_only)
            GoTo xt
        End If
    End With
        
xt: Set fso = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function WbkIsOpen(ByVal wb As Variant, _
                  Optional ByRef wb_result As Workbook) As Boolean
' -----------------------------------------------------------------------------------
' Returns TRUE when the Workbook (wb) - which may be a Workbook object, a Workbook's
' name or fullname - is open in any Excel Application instance. If a fullname is
' provided and the file does not exist under this full name but a Workbook with the
' given name is open (but from another folder) the Workbook is regarded moved and
' thus is returned as open object(wb_result).
' Because Workbooks with the same WbName may be open when they have different
' extensions a Workbook's Name including its extension is checked.
' -----------------------------------------------------------------------------------
    Const PROC = "WbkIsOpen"
    
    On Error GoTo eh
    Dim OpenWbks As Dictionary
    Dim OpenWbk  As Workbook
    Dim fso      As New FileSystemObject
    Dim WbName As String
    
    If Not mWbk.IsWbObject(wb) And Not mWbk.IsFullName(wb) And Not mWbk.IsName(wb) And Not TypeName(wb) = "String" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter wb) is neither a Workbook object nor a Workbook's name or fullname)!"
       
    Set OpenWbks = mWbk.Opened
    If mWbk.IsName(wb) Then
        '~~ wb is a Workbook's Name including its extension
        WbName = fso.GetFileName(wb)
        If OpenWbks.Exists(WbName) Then
            '~~ A Workbook with the same 'WbName' is open
            Set OpenWbk = OpenWbks.Item(WbName)
            '~~ When a Workbook's Name is provided the Workbook is only regarde open when the open
            '~~ Workbook has the same name (i.e. including its extension)
            If fso.GetFile(OpenWbk.FullName).name <> fso.GetFileName(wb) Then Set OpenWbk = Nothing
        End If
    ElseIf mWbk.IsFullName(wb) Then
        WbName = fso.GetFileName(wb)
        If OpenWbks.Exists(WbName) Then
            '~~ A Workbook with the same 'WbName' is open
            Set OpenWbk = OpenWbks.Item(WbName)
            '~~ The provided (wb) specifies an exist Workbook file. This Workbook is regarded open (and returned as opject)
            '~~ when a Workbook with its Name (including the extension!) is open regardless in which location
            If fso.GetFile(OpenWbk.FullName).name <> fso.GetFileName(wb) Then Set OpenWbk = Nothing
        End If
    ElseIf mWbk.IsWbObject(wb) Then
        WbName = wb.name
        If Opened.Exists(WbName) Then
            Set OpenWbk = OpenWbks.Item(WbName)
        End If
    Else
        '~~ If wb is a Workbook's WbName it is regarded open when one with that WbName is open
        '~~ regrdless its extension
        If OpenWbks.Exists(wb) Then Set OpenWbk = OpenWbks.Item(wb)
    End If
    
xt: If mWbk.IsWbObject(OpenWbk) Then
        WbkIsOpen = True
        Set wb_result = OpenWbk
    End If
    Set fso = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Function URLExists(ByVal ue_url As String) As Boolean
' -----------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------
    Dim oRequest    As Object
    Dim var         As Variant

    On Error GoTo xt
    Set oRequest = CreateObject("WinHttp.WinHttpRequest.5.1")

    With oRequest
      .Open "GET", ue_url, False
      .send
      var = .StatusText
    End With
    URLExists = var = "OK"

xt: Set oRequest = Nothing
    Exit Function

End Function

