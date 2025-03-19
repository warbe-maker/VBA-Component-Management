VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Class Module clsCode: Provides methods and properties for the code of/in a
' ===================== VB-CodeModule or an Export-File. The class is used
' as two instances, one (ProcOrigin) representing a procedures origin code
' and a second (ProcCopy) representing a procedures code copy used in any
' other component but the origin component.
'
' Public Methods:
' ----------------
' AsFile        Returns the mere code of a source as file, i.e. any leading
'               or trailing empty code lines excluded. An "exclude empty code
'               lines" option may be used to exclude any empty code line.
' DiffersFrom   Returns TRUE when the code differs from the code of another
'               source.
' DsplyDiffs    Displays the difference of the code of two sources by means
'               of WinMerge.
' ReplaceWith   Replaces the code in the provided CodeModule's procedure
'               with the code provided by another instance of this Class
'               Module.
'
' Public Properties:
' ------------------
' AsString    r Provides the code as a single string with the code lines
'               delimited by a vbCrLf
' CompName    r Name of the component for which the code is provided.
' IsNone      r Returns TRUE when the provided source did not contain any
'               code lines or none but empty lines.
' Source      w The source from which the code is to be provided, either
'               an Export-File as object or full name or a VB-Project
'               CodeModule.
'
' W. Rauschenberger, Berlin May 2024
' ---------------------------------------------------------------------------
Private Const ERR_NO_SOURCE = "Obtaining the code as array or string when the source 'is none' is useless!" & vbLf & _
                              "The Property 'IsNone' may be used to avoid this. The Property is TRUE when " & _
                              "there are no known code lines for whichever reason, no source ever provided, " & _
                              "the source does not contain any code lines, or just empty code lines)."
Private sAsTrimmedString    As String   ' Code as a single trimmed string (without leading and trailing empty code lines)
Private vSource             As Variant
Private sSourceName         As String
Private sCompName           As String
Private sSourceDesc         As String
Private bSourceIsNone       As Boolean
Private dctTempFiles        As New Dictionary
Private sCodeAsString       As String

Private Sub Class_Initialize()
    bSourceIsNone = True
End Sub

Private Sub Class_Terminate()
    FileTempRemove
End Sub

Private Property Get Arry(Optional ByRef c_arr As Variant, _
                          Optional ByVal c_index As Long = -1) As Variant
' ----------------------------------------------------------------------------
' Universal 1-dimensional array read procedure. Returns Empty when a given
' array (c_arr) is not allocated or a provided index is beyond/outside current
' dimensions.
' ----------------------------------------------------------------------------
    Dim i As Long
    
    Arry = Empty ' the default
    If IsArray(c_arr) Then
        On Error Resume Next
        i = LBound(c_arr)
        If Err.Number = 0 Then
            If c_index >= LBound(c_arr) And c_index <= UBound(c_arr) _
            Then Arry = c_arr(c_index)
        End If
    End If
    
End Property

Private Property Let Arry(Optional ByRef c_arr As Variant, _
                          Optional ByVal c_index As Long = -99, _
                                   ByVal c_var As Variant)
' ----------------------------------------------------------------------------
' Universal 1-dimensional array add/update procedure, avoiding any prior
' checks whether allocated, empty not yet existing, etc.
' - Adds an item (c_var) to an array (c_arr) when no index is provided or when
'   the index lies beyond UBound
' - When an index is provided, the item is inserted/updated at the given
'   index - even when the array yet doesn't exist or is not yet allocated.
' ----------------------------------------------------------------------------
    Const PROC = "Arry(Let)"
        
    On Error GoTo eh
    If ArryIsAllocated(c_arr) Then
        '~~ The array has at least one item
        Select Case True
            Case c_index = -99                              ' When for an allocated array no index is provided, the item is added
                ReDim Preserve c_arr(UBound(c_arr) + 1)
                c_arr(UBound(c_arr)) = c_var
            Case c_index >= 0 And c_index <= UBound(c_arr)  ' Replace an existing item
                c_arr(c_index) = c_var
            Case c_index > UBound(c_arr)                    ' New item beyond current UBound
                ReDim Preserve c_arr(c_index)
                c_arr(c_index) = c_var
            Case c_index < LBound(c_arr)
                Err.Raise AppErr(2), ErrSrc(PROC), "Index is less than LBound of array!"
        End Select
        
    Else
        '~~ The array does yet not exist
        If c_index = -99 Then
            '~~ When no index is provided the item is the first of a new array
            c_arr = Array(c_var)
        ElseIf c_index >= 0 Then
            ReDim Preserve c_arr(c_index)
            c_arr(c_index) = c_var
        Else
            Err.Raise AppErr(3), ErrSrc(PROC), "the provided index is less than 0!"
        End If
    End If
    
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get CompName() As String:           CompName = sCompName:   End Property

Public Property Let CompName(ByVal s As String):    sCompName = s:          End Property

Public Property Get IsNone() As Boolean:            IsNone = bSourceIsNone: End Property

Public Property Let Source(ByVal s_src As Variant)
' ----------------------------------------------------------------------------
' Depending on the provided argument's (s_src) type the code is saved as an
' array of code lines whereby leading and trailing empty code lines are
' excempted. Types acepted are: - String, considered a file's full name
'                               - File, considered a VBComponent's Export-File
'                               - VBComponent
'                               - CodeModule
' Note: The module is initialized with IsNone = TRUE and this is only set to
'       FALSE when a Source had been provided.
' ----------------------------------------------------------------------------
    
    Select Case TypeName(s_src)
        Case "String"
            If s_src <> vbNullString Then
                With New FileSystemObject
                    If .FileExists(s_src) Then
                        vSource = s_src
                        sSourceName = s_src
                        bSourceIsNone = False
                    End If
                End With
            End If
        Case "VBComponent": Set vSource = s_src
                            sSourceName = s_src.Name
                            bSourceIsNone = False
        Case "CodeModule":  Set vSource = s_src
                            sSourceName = s_src.Parent.Name
                            bSourceIsNone = False
        Case "File":        Set vSource = s_src
                            sSourceName = s_src.Path
                            bSourceIsNone = False
    End Select
    If Not bSourceIsNone Then sCodeAsString = vbNullString
    
End Property

Public Property Get SourceDesc() As String:         SourceDesc = sSourceDesc:   End Property

Public Property Let SourceDesc(ByVal s As String):  sSourceDesc = s:            End Property

Public Property Get SourceName() As String:         SourceName = sSourceName:   End Property

Private Function ArrayIsAllocated(a_v As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = UBound(a_v) >= LBound(a_v)
    On Error GoTo 0
    Err.Clear
End Function

Public Function AsFile(ByVal a_name As String, _
              Optional ByVal a_empty_excluded As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns the code from the Source as temporary file with a given name
' (a_name), optionally with empty code lines excluded.
' ----------------------------------------------------------------------------
    Dim sTempFile As String
    
    sTempFile = TempFile(a_name)
    With New FileSystemObject
        If .FileExists(sTempFile) _
        Then .DeleteFile sTempFile
    End With
    
    If Not a_empty_excluded _
    Then StringAsFile sAsTrimmedString, sTempFile _
    Else StringAsFile CodeAsString, sTempFile
    
    Set AsFile = fso.GetFile(sTempFile)
    
End Function

Private Function AsTrimmedArray(Optional ByVal a_empty_excluded As Boolean = False) As Variant
' ----------------------------------------------------------------------------
' Returns th code as a trimmed array, i.e. without any leading and trailing
' empty code lines.
' ----------------------------------------------------------------------------
    Const PROC = "AsTrimmedArray"
    
    If bSourceIsNone Then Err.Raise AppErr(1), ErrSrc(PROC), ERR_NO_SOURCE
    If Not a_empty_excluded _
    Then AsTrimmedArray = Split(sAsTrimmedString, vbCrLf) _
    Else AsTrimmedArray = Split(CodeAsString, vbCrLf)
    
End Function

Public Function AsTrimmedString(Optional ByVal a_empty_excluded As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the code as a single string only without leading and trailing empty
' code lines.
' ----------------------------------------------------------------------------
    Const PROC = "AsTrimmedString"
    If bSourceIsNone Then Err.Raise AppErr(1), ErrSrc(PROC), ERR_NO_SOURCE
    
    If Not a_empty_excluded Then
        AsTrimmedString = StringTrimmed(sAsTrimmedString)
    Else
        AsTrimmedString = CodeAsString
    End If
    
End Function

Public Function CodeAsString() As String
' ----------------------------------------------------------------------------
' Returns the code of the provided source, which is an Export-File, a
' VBComponent or a CodeModule.
' ----------------------------------------------------------------------------
    Dim fl      As File
    Dim vbcm    As CodeModule
    
    If sCodeAsString = vbNullString Then
        Select Case TypeName(vSource)
            Case "String"
                If vSource <> vbNullString Then
                    With fso
                        If .FileExists(vSource) Then
                            Set fl = .GetFile(vSource)
                            sCodeAsString = StringTrimmed(StringFromExpFile(fl))
                        End If
                    End With
                End If
            Case "VBComponent": Set vbcm = vSource.CodeModule
                                If vbcm.CountOfLines > 0 Then sCodeAsString = StringFromComp(vbcm)
            Case "CodeModule":  Set vbcm = vSource
                                If vbcm.CountOfLines > 0 Then sCodeAsString = StringFromComp(vbcm)
            Case "File":        Set fl = vSource
                                sCodeAsString = StringTrimmed(StringFromExpFile(vSource))
        End Select
        If sCodeAsString <> vbNullString Then
            sCodeAsString = StringEmptyExcluded(sCodeAsString)
        End If
    End If
    CodeAsString = sCodeAsString
    
End Function

Public Function DiffersFrom(ByVal d_from_code As clsCode, _
                   Optional ByRef d_diffs As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when this class instance's code (AsTrimmedString) differs from
' those of another's instance (d_from_code). The function returns the first
' 5 different code lines as array in the form <line-number> <source> <line>
' i.e each difference is returnes as two items.
' ----------------------------------------------------------------------------
    Const PROC = "DiffersFrom"
    
    On Error GoTo eh
    Dim arr As Variant
    
    Select Case True
        Case Not IsNone And Not d_from_code.IsNone
            DiffersFrom = StrComp(CodeAsString _
                                , d_from_code.CodeAsString _
                                , vbTextCompare) <> 0
            If DiffersFrom Then
                DiffersFromDiffs d_from_code, arr
            End If
            
        Case Not Me.IsNone And d_from_code.IsNone:   DiffersFrom = True
        Case Me.IsNone And Not d_from_code.IsNone:   DiffersFrom = True
        Case Me.IsNone And d_from_code.IsNone
    End Select
    
xt: d_diffs = arr
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub DiffersFromDiffs(ByVal d_from_code As clsCode, _
                   Optional ByRef d_diffs As Variant = Nothing, _
                   Optional ByVal d_diffs_limit As Long = 1)
' ----------------------------------------------------------------------------
' Provides an array (d_diffs) with n (d_diffs_limit) code differences, each
' as a pair of different lines.
' ----------------------------------------------------------------------------
    Const PROC = "DiffersFromDiffs"
    
    On Error GoTo eh
    Dim i       As Long
    Dim arrThis As Variant
    Dim arrFrom As Variant
    Dim arrDiff As Variant
    Dim j       As Long
    Dim lMaxSrc As Long
    
    lMaxSrc = Max(Len(sSourceDesc), Len(d_from_code.SourceDesc)) + 4
    arrThis = Split(CodeAsString, vbCrLf)
    arrFrom = Split(d_from_code.CodeAsString, vbCrLf)
    For i = 1 To Min(UBound(arrThis), UBound(arrFrom))
        If StrComp(arrThis(i), arrFrom(i), vbTextCompare) <> 0 Then
            Arry(arrDiff) = "Line " & Format(i, "0000") & " " & mBasic.AlignLeft(sSourceDesc, lMaxSrc, ".:") & " " & arrThis(i)
            Arry(arrDiff) = "Line " & Format(i, "0000") & " " & mBasic.AlignLeft(d_from_code.SourceDesc, lMaxSrc, ".:") & " " & arrFrom(i)
            j = j + 1
            If j = d_diffs_limit Or i >= Min(UBound(arrThis), UBound(arrFrom)) Then Exit For
        End If
    Next i
    d_diffs = arrDiff
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub DsplyDiffs(ByVal d_left_this_file_name As String, _
                      ByVal d_left_this_file_title As String, _
                      ByVal d_rght_vrss_code As clsCode, _
                      ByVal d_rght_vrss_file_name As String, _
                      ByVal d_rght_vrss_file_title As String)
' ----------------------------------------------------------------------------
' Left is the current code, Right is the changed code
' ----------------------------------------------------------------------------
    Const PROC As String = "DsplyDiffs"
    
    On Error GoTo eh
    Dim sFileThis   As String
    Dim sFileFrom   As String
    
    sFileThis = TempFile(d_left_this_file_name)
    sFileFrom = TempFile(d_rght_vrss_file_name)
    
    StringAsFile d_rght_vrss_code.CodeAsString, sFileFrom
    StringAsFile CodeAsString, sFileThis
    
    DsplyFilesDiff d_file_left_full_name:=sFileThis _
                 , d_file_right_full_name:=sFileFrom _
                 , d_file_left_title:=d_left_this_file_title _
                 , d_file_right_title:=d_rght_vrss_file_title
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function DsplyFilesDiff(ByVal d_file_left_full_name As String, _
                                ByVal d_file_right_full_name As String, _
                                ByVal d_file_left_title As String, _
                                ByVal d_file_right_title As String) As Boolean
' ----------------------------------------------------------------------------
' Displays the differences between export file 1 and 2 by means of WinMerge!
' Note: CompMan ignores any differences caused by empty code lines or case.
'       When a difference is displayed it is thus not because of this kind of
'       differneces but of others. Unfortunately it depends on the installed
'       WinMerge's set option wether or not these kind of differences are
'       displayed.
' ----------------------------------------------------------------------------
    Const PROC              As String = "ExpFilesDiffDisplay"
    
    On Error GoTo eh
    Dim waitOnReturn        As Boolean: waitOnReturn = True
    Dim windowStyle         As Integer: windowStyle = 1
    Dim sCommand            As String
    Dim wshShell            As Object
    
    If Not AppIsInstalled("WinMerge") _
    Then Err.Raise Number:=AppErr(1) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="WinMerge is obligatory for the Compare service of this module but not installed!" & vbLf & vbLf & _
                                "(See ""https://winmerge.org/downloads/?lang=en"" for download)"
        
    If Not fso.FileExists(d_file_left_full_name) _
    Then Err.Raise Number:=AppErr(2) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & d_file_left_full_name & """ does not exist!"
    
    If Not fso.FileExists(d_file_right_full_name) _
    Then Err.Raise Number:=AppErr(3) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & d_file_right_full_name & """ does not exist!"
        
    '~~ Guarantee an ini-File which meets CompMan's specifics
    With fso
        If .FileExists(mWinMergeIni.WinMergeIniFullName) Then .DeleteFile (mWinMergeIni.WinMergeIniFullName)
        mWinMergeIni.Setup
    End With
    
    '~~ Prepare command line
    sCommand = "WinMergeU " & _
               """" & d_file_left_full_name & """" & " " & _
               """" & d_file_right_full_name & """" & _
               " /e " & _
               " /dl " & DQUOTE & d_file_left_title & DQUOTE & _
               " /dr " & DQUOTE & d_file_right_title & DQUOTE & " " & _
               " /inifile " & """" & mWinMergeIni.WinMergeIniFullName & """"

    '~~ Execute command line
    Set wshShell = CreateObject("WScript.Shell")
    With wshShell
        DsplyFilesDiff = .Run(Command:=sCommand, windowStyle:=windowStyle, waitOnReturn:=waitOnReturn)
    End With
        
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCode" & "." & sProc
End Function

Private Sub FileTempRemove()
' ----------------------------------------------------------------------------
' Removes all temporary files created.
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    With fso
        For Each v In dctTempFiles
            If .FileExists(v) _
            Then .DeleteFile v
        Next v
    End With
    
End Sub

Public Function Meets(ByVal d_other_code As clsCode) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when this code meets the other instance's code (d_other_code).
' Note: Error is passed on to the caller.
' ----------------------------------------------------------------------------
    Const PROC = "Meets"
    
'    mBasic.BoP ErrSrc(PROC)
    If Not IsNone Then
        If Not d_other_code.IsNone Then
            If StrComp(CodeAsString _
                     , d_other_code.CodeAsString _
                     , vbTextCompare) = 0 Then
                Meets = True
            End If
        Else
            '~~ False is returned when the other code is none
        End If
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), "Code meets code failed since this code is none!"
    End If
'    mBasic.EoP ErrSrc(PROC)
    
End Function

Public Sub ReplaceWith(ByVal r_source_code As clsCode)
' ----------------------------------------------------------------------------
' Replaces the code lines - if any - in the provided CodeModule (r_this_vbcm)
' with the code of the provided source (r_source_code).
' ----------------------------------------------------------------------------
    Const PROC = "ReplaceWith"
    
    Dim vbcm As CodeModule
    
    If Not TypeName(vSource) = "CodeModule" _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Unable to replace the code because the source is not a CodeModule!"
    
    Set vbcm = vSource
    With vbcm
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .InsertLines 1, r_source_code.AsTrimmedString
    End With

End Sub

Private Function SplitString(ByVal s_s As String) As String
    
    Select Case True
        Case InStr(s_s, vbCrLf) <> 0: SplitString = vbCrLf
        Case InStr(s_s, vbCr) <> 0:   SplitString = vbCr
        Case InStr(s_s, vbLf) <> 0:   SplitString = vbLf
    End Select
    If Len(SplitString) = 0 Then SplitString = vbCrLf
    
End Function

Private Function StringAsFile(ByVal s_strng As String, _
                     Optional ByRef s_file As Variant = vbNullString, _
                     Optional ByVal s_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes a string (s_strng) to a file (s_file) which might be a file object or
' a file's full name. When no file (s_file) is provided, a temporary file is
' returned.
' Note 1: Only when the string has sub-strings delimited by vbCrLf the string
'         is written a records/lines.
' Note 2: When the string has the alternate split indicator "|&|" this one is
'         replaced by vbCrLf.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim sSplit As String
    
    sSplit = SplitIndctr(s_strng)
    If sSplit <> vbCrLf Then s_strng = Replace(s_strng, sSplit, vbCrLf)
    
    Select Case True
        Case s_file = vbNullString: s_file = TempFile
        Case TypeName(s_file) = "File": s_file = s_file.Path
    End Select
    
    If s_file_append _
    Then Open s_file For Append As #1 _
    Else Open s_file For Output As #1
    Print #1, s_strng
    Close #1
    Set StringAsFile = fso.GetFile(s_file)
    
End Function

Public Function StringEmptyExcluded(ByVal s_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' ----------------------------------------------------------------------------
    Dim sSplit As String
    
    '~~ Get line terminating string
    sSplit = SplitString(s_s)
    
    s_s = StringTrimmed(s_s) ' leading and trailing empty already excluded
    
    Do While InStr(s_s, sSplit & sSplit) <> 0
        s_s = Replace(s_s, sSplit & sSplit, sSplit)
    Loop
    StringEmptyExcluded = s_s
    
End Function

Private Function StringFromComp(ByVal s_vbcm As CodeModule) As Variant
' ----------------------------------------------------------------------------
' Returns a string with all the code lines in the CodeModule (s_vbcm) whereby
' any leading and trailing code lines are excluded.
' ----------------------------------------------------------------------------
    Dim sCodeLines  As String

    With s_vbcm
        If .CountOfLines > 0 Then
            Set vSource = s_vbcm
            sCodeLines = .Lines(1, .CountOfLines)
            '~~ Eliminate leading empty code lines
            Do While Left(sCodeLines, 2) = vbCrLf
                sCodeLines = Right(sCodeLines, Len(sCodeLines) - 2)
            Loop
            '~~ Eliminate trailing empty code lines
            Do While Right(sCodeLines, 2) = vbCrLf
                sCodeLines = Left(sCodeLines, Len(sCodeLines) - 2)
            Loop
            If Len(sCodeLines) >= 1 Then
                '~~ The component not only contains empty code lines
                bSourceIsNone = False
                StringFromComp = sCodeLines
            End If
        End If
    End With

End Function

Private Function StringFromExpFile(ByVal a_fl As File) As String
' ----------------------------------------------------------------------------
' Returns an array with the code lines in the Export-File (a_fl) whereby
' any empty leading and trailing code lines and any code line beginning with
' "Attribute " is excempted.
'
' Note: Disregarding the used line terminating charcter/string in the input
'       file (a_fl), the result line terminating string will be vbCrLf.
' ----------------------------------------------------------------------------
    Const PROC = "StringFromExpFile"
    
    On Error GoTo eh
    Dim aLinesExpFile   As Variant
    Dim aLinesMereCode  As Variant
    Dim i               As Long
    Dim sLinesExpFile   As String
    Dim sSplit          As String
    Dim lFirstCodeLine  As Long
    Dim lCodeLines      As Long
    Dim j               As Long
    
    Open a_fl For Input As #1
    sLinesExpFile = Input$(LOF(1), 1)
    Close #1
    
    '~~ Get line terminating string
    sSplit = SplitString(sLinesExpFile)
    
    '~~ Eliminate trailing empty lines
    Do While Right(sLinesExpFile, 4) = sSplit & sSplit
        sLinesExpFile = Left(sLinesExpFile, Len(sLinesExpFile) - 2)
    Loop
    
    aLinesExpFile = Split(sLinesExpFile, sSplit)
    '~~ Skip all lines before any "Attribute VB_...." line
    '~~ Note that this assumes that an export file has at least one "Attribute " line
    '~~ possibly without any preceeding lines!
    Do While Not aLinesExpFile(i) Like "Attribute*"
        If i = UBound(aLinesExpFile) Then Exit Do
        i = i + 1
    Loop
        
    '~~ Move all code lines disregarding any empty lines but excluding any code line beginning "Attribute ".
    '~~ Note! An "Attribute " code line may appear also "among" other code lines.
    
    ReDim aLinesMereCode(UBound(aLinesExpFile))
    j = 0
    For i = i To UBound(aLinesExpFile)
        If Not aLinesExpFile(i) Like "Attribute *" Then
            aLinesMereCode(j) = aLinesExpFile(i)
            j = j + 1
        End If
    Next i
    ReDim Preserve aLinesMereCode(j - 1)
    
    If ArrayIsAllocated(aLinesMereCode) _
    Then StringFromExpFile = Join(aLinesMereCode, vbCrLf) _
    Else StringFromExpFile = vbNullString
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function StringTrimmed(ByVal s_s As String, _
                      Optional ByRef s_as_dict As Dictionary = Nothing) As String
' ----------------------------------------------------------------------------
' Returns the code (s_s) provided as a single string without leading and
' trailing empty lines. When a Dictionary is provided the string is returned
' as items with the line number as key.
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim i       As Long
    Dim v       As Variant
    
    s = s_s
    '~~ Eliminate leading empty code lines
    Do While Left(s, 2) = vbCrLf
        s = Right(s, Len(s) - 2)
    Loop
    '~~ Eliminate trailing eof
    If Right(s, 1) = VBA.Chr(26) Then
        s = Left(s, Len(s) - 1)
    End If
    
    '~~ Eliminate trailing empty code lines
    Do While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Loop
    Do While Right(s, 2) = vbCr & vbLf
        s = Left(s, Len(s) - 2)
    Loop
    
    Do While Right(s, 1) = vbCr Or Right(s, 1) = vbLf
        s = Left(s, Len(s) - 1)
    Loop
    
    If Right(s, 2) = vbCrLf Then Stop
    If Right(s, 1) = vbCr Then Stop
    If Right(s, 1) = vbLf Then Stop
    
    StringTrimmed = Trim(s)
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

Private Function TempFile(Optional ByVal f_name As String = vbNullString, _
                          Optional ByVal f_path As String = vbNullString, _
                          Optional ByVal f_extension As String = ".txt") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary file, randomly named when no name
' (f_name) is provided, in the system temp folder when no path is provided, with
' an extension .txt when none is provided.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    If f_name = vbNullString Then
        sTemp = Replace(fso.GetTempName, ".tmp", f_extension)
    Else
        sTemp = f_name & f_extension
    End If
    If f_path = vbNullString Then f_path = fso.GetSpecialFolder(2)
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp

End Function

