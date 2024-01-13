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

Private Const NO_SOURCE As String = "Obtaining the code as array or string when the source 'is none' is useless!" & vbLf & _
                                    "The Property 'IsNone' may be used to avoid this. The Property is TRUE when " & _
                                    "there are no known code lines for whichever reason, no source ever provided, " & _
                                    "the source does not contain any code lines, or just empty code lines)."
Private aAsArray                As Variant  ' Code as a trimmed array (without leading and trailing empty code lines)
Private aAsArrayEmptyExcluded    As Variant  ' Code as array with any empty code lines excluded
Private sAsString               As String   ' Code as a single timmed string (without leading and trailing empty code lines)
Private sAsStringEmptyExcluded   As String   ' Code as a single timmed string (without leading and trailing empty code lines)
Private vSource                 As Variant
Private sCompName               As String
Private bSourceIsNone           As Boolean
Private flCode                  As File

Public Property Get AsArrayEmptyExcluded() As Variant
' ----------------------------------------------------------------------------
' Code as a array without leading and trailing empty code lines.
' ----------------------------------------------------------------------------
    Const PROC = "AsArray/Get"
    Dim s As String
    
    If bSourceIsNone Then Err.Raise AppErr(1), ErrSrc(PROC), NO_SOURCE
    If Not ArrayIsAllocated(aAsArrayEmptyExcluded) Then
        s = Join(aAsArray, vbCrLf)
        
    End If
    AsArrayEmptyExcluded = aAsArrayEmptyExcluded

End Property

Public Property Get AsArray() As Variant
' ----------------------------------------------------------------------------
' Code as a array without leading and trailing empty code lines.
' ----------------------------------------------------------------------------
    Const PROC = "AsArray/Get"
    If bSourceIsNone Then Err.Raise AppErr(1), ErrSrc(PROC), NO_SOURCE
    AsArray = aAsArray
    
End Property

Public Property Get AsStringEmptyExcluded() As String
' ----------------------------------------------------------------------------
' Returns the code as a single string with any empty code line excluded.
' ----------------------------------------------------------------------------
    Const PROC = "AsStringEmptyExcluded"
    
    Dim i As Long
    Dim j As Long
    Dim s As String
    
    If bSourceIsNone Then Err.Raise AppErr(1), ErrSrc(PROC), NO_SOURCE
    If sAsStringEmptyExcluded = vbNullString Then
        sAsStringEmptyExcluded = StringEmptyExcluded(AsString)
    End If
    AsStringEmptyExcluded = sAsStringEmptyExcluded
    
End Property

Private Function StringEmptyExcluded(ByVal s_s As String) As String
    Dim i   As Long
    
    i = InStr(s_s, vbCrLf & vbCrLf)
    Do While i <> 0
        s_s = Replace(s_s, vbCrLf & vbCrLf, vbCrLf)
        i = InStr(s_s, vbCrLf & vbCrLf)
    Loop
    StringEmptyExcluded = s_s
    
End Function

Public Property Get AsString() As String
' ----------------------------------------------------------------------------
' Returns the code as a single string only without leading and trailing empty
' code lines.
' ----------------------------------------------------------------------------
    Const PROC = "AsString/Get"
    If bSourceIsNone Then Err.Raise AppErr(1), ErrSrc(PROC), NO_SOURCE
    If sAsString = vbNullString Then
        sAsString = Join(aAsArray, vbCrLf)
    End If
    AsString = sAsString
    
End Property

Public Property Get CompName() As String:           CompName = sCompName:   End Property

Public Property Let CompName(ByVal s As String):    sCompName = s:          End Property

Public Property Get IsNone() As Boolean:            IsNone = bSourceIsNone: End Property

Public Property Let Source(ByVal f_v As Variant)
' ----------------------------------------------------------------------------
' Depending on the provided argument's (f_v) type the code is saved as an
' array of code lines whereby leading and trailing empty code lines are
' excempted. Types acepted are: - String, considered a file's full name
'                               - File, considered a VBComponent's Export-File
'                               - VBComponent
'                               - CodeModule
' Note: The module is initialized with IsNone = TRUE and this is only set to
'       FALSE when a Source had been provided which contains at least one non
'       non empty code line.
' ----------------------------------------------------------------------------
    Const PROC = "Source/Let"
    
    Dim fl      As File
    Dim vbc     As VBComponent
    Dim vbcm    As CodeModule
    
    Select Case TypeName(f_v)
        Case "String"
            If f_v <> vbNullString Then
                With New FileSystemObject
                    If .FileExists(f_v) Then
                        vSource = f_v
                        sCompName = fso.GetBaseName(f_v)
                        aAsArray = ArrayFromExpFile(.GetFile(f_v))
                    End If
                End With
            End If
        Case "VBComponent"
            Set vbc = f_v
            Set vbcm = vbc.CodeModule
            If vbcm.CountOfLines > 0 Then
                aAsArray = ArrayFromComp(vbcm)
            Else
                bSourceIsNone = True
            End If
        Case "CodeModule"
            Set vbcm = f_v
            sCompName = vbcm.Parent.Name
            If vbcm.CountOfLines > 0 Then
                aAsArray = ArrayFromComp(vbcm)
            Else
                bSourceIsNone = True
            End If
        Case "File"
            Set fl = f_v
            Set vSource = f_v
            sCompName = fso.GetBaseName(fl.Path)
            aAsArray = ArrayFromExpFile(fl)
        Case Else
            bSourceIsNone = True
    End Select
    
End Property

Private Function ArrayFromComp(ByVal c_vbcm As CodeModule) As Variant
' ----------------------------------------------------------------------------
' Returns an array with the code lines in the CodeModule (c_cbcm) whereby
' any empty leading and trailing code lines are already excempted.
' ----------------------------------------------------------------------------

    Dim iEmpty              As Long
    Dim sCode               As String ' code without any leading and trailing empty code lines
    
    With c_vbcm
        If .CountOfLines > 0 Then
            Set vSource = c_vbcm
            sCode = .Lines(1, .CountOfLines)
            '~~ Eliminate leading empty code lines
            Do While Left(sCode, 2) = vbCrLf
                sCode = Right(sCode, Len(sCode) - 2)
            Loop
            '~~ Eliminate trailing empty code lines
            Do While Right(sCode, 2) = vbCrLf
                sCode = Left(sCode, Len(sCode) - 2)
            Loop
            If Len(sCode) >= 1 Then
                '~~ The component not only contains empty code lines
                sAsString = sCode ' already done on the fly
                aAsArray = ArrayTrimmed(Split(sCode, vbCrLf))
                bSourceIsNone = False
                '~~ Eliminate any empty code lines
                sAsStringEmptyExcluded = StringEmptyExcluded(sCode)
                aAsArrayEmptyExcluded = Split(sAsStringEmptyExcluded, vbCrLf)
            End If
        End If
    End With
    
End Function

Private Function ArrayFromExpFile(ByVal a_fl As File) As Variant
' ----------------------------------------------------------------------------
' Returns an array with the code lines in the Export-File (a_fl) whereby
' any empty leading and trailing code lines are already excempted.
' ----------------------------------------------------------------------------
    Const PROC = "ArrayFromExpFile"
    
    On Error GoTo eh
    Dim sCode       As String
    Dim aLines()    As String
    Dim i           As Integer
    Dim j           As Integer
    Dim arr         As Variant
    
    Open a_fl For Input As #1
    sCode = Input$(lOf(1), 1)
    Close #1
    
    aLines = Split(sCode, vbCrLf)
    '~~ Skip all lines before any "Attribute VB_...." line
    Do While Not aLines(i) Like "Attribute VB_*"
        i = i + 1
        If i > UBound(aLines) Then Exit Do
    Loop
    
    '~~ Skip all "Attribute VB_...." lines
    Do Until Not aLines(i) Like "Attribute VB_*"
        i = i + 1
        If i > UBound(aLines) Then Exit Do
    Loop
    
    '~~ Move all remaining code lines up - disregarding any empty lines
    j = 0
    For i = i To UBound(aLines)
        aLines(j) = aLines(i)
        j = j + 1
    Next i
    
    '~~ Redim the array to the - now shortend content
    If ArrayIsAllocated(aLines) And j <> 0 Then
        ReDim Preserve aLines(j - 1)
        arr = ArrayTrimmed(aLines)
        bSourceIsNone = False
    End If
    ArrayFromExpFile = arr
    sAsString = Join(arr, vbCrLf)
    sAsStringEmptyExcluded = StringEmptyExcluded(sAsString)
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Function ArrayIsAllocated(a_v As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = Not IsError(UBound(a_v))
End Function

Private Function ArrayTrimmed(ByVal t_v As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "ArrayTrimmed"
    
    Dim i   As Long
    Dim j   As Long
    Dim k   As Long
    Dim arr As Variant
    
    If Not ArrayIsAllocated(t_v) Then
        bSourceIsNone = True
        Exit Function
    End If
    
    '~~ Get first code line not empty
    For i = LBound(t_v) To UBound(t_v)
        If Len(Trim(t_v(i))) > 0 Then Exit For
    Next i
    
    '~~ Get last code line not empty
    For j = UBound(t_v) To LBound(t_v) Step -1
        If Len(Trim(t_v(j))) > 0 Then Exit For
    Next j
    
    If i < 0 Or j < 0 Then
        bSourceIsNone = True
    Else
        ReDim arr(j)
        For i = i To j
            arr(k) = t_v(i)
            k = k + 1
        Next i
        ArrayTrimmed = arr
        bSourceIsNone = False
    End If
    
End Function

Public Function AsFile(ByVal a_file_name As String) As File
' ----------------------------------------------------------------------------
' Returns the code in a VBcomponent or an Export-File as a File with leading
' and trailing empty code lines eliminated. The function is used to display
' a code difference.
' ----------------------------------------------------------------------------
    Set AsFile = FileFromArray(aAsArray, a_file_name)
End Function

Public Function DiffersFrom(ByVal d_from As clsCode, _
                   Optional ByVal d_empty_included As Boolean = True) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when this code lines differ from those (d_from) code lines
' ----------------------------------------------------------------------------
    Const PROC = "DiffersFrom"
    
    On Error GoTo eh
    Dim i       As Long
    Dim arr     As Variant
    Dim arrFrom As Variant
    
    Select Case True
        Case Not IsNone And Not d_from.IsNone
            If d_empty_included Then
                DiffersFrom = StrComp(AsString, d_from.AsString, vbTextCompare)
                Debug.Print "changed (empty included): " & sCompName
                If DiffersFrom Then
                    Debug.Print "changed (empty included): " & sCompName & vbLf & "The first detected difference:"
                    arr = Split(AsString, vbCrLf)
                    arrFrom = Split(d_from.AsString, vbCrLf)
                    For i = 1 To Min(UBound(arr), UBound(arrFrom))
                        If StrComp(arr(i), arrFrom(i), vbTextCompare) <> 0 Then
                            Debug.Print "Line " & i & """" & arr(i) & """"
                            Debug.Print "Line " & i & """ & arrFrom(i) & """""
                            Exit For
                        End If
                    Next i
                End If
            Else
                DiffersFrom = StrComp(AsStringEmptyExcluded, d_from.AsStringEmptyExcluded, vbTextCompare)
                If DiffersFrom Then
                    Debug.Print "changed (empty excluded): " & sCompName & vbLf & "The first detected difference:"
                    arr = Split(AsStringEmptyExcluded, vbCrLf)
                    arrFrom = Split(d_from.AsStringEmptyExcluded, vbCrLf)
                    For i = 1 To Min(UBound(arr), UBound(arrFrom))
                        If StrComp(arr(i), arrFrom(i), vbTextCompare) <> 0 Then
                            Debug.Print "Line " & i & """" & arr(i) & """"
                            Debug.Print "Line " & i & """ & arrFrom(i) & """""
                            Exit For
                        End If
                    Next i
                End If
            End If
            
        Case Not Me.IsNone And d_from.IsNone
            DiffersFrom = True
        Case Me.IsNone And Not d_from.IsNone
            DiffersFrom = True
        Case Me.IsNone And d_from.IsNone
            Err.Raise AppErr(1), ErrSrc(PROC), NO_SOURCE
    End Select
    
    If DiffersFrom Then
    Else
        Debug.Print "unchanged: " & sCompName
    End If
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub DiffFromDsply(ByVal d_this_file_name As String, _
                         ByVal d_this_file_title As String, _
                         ByVal d_from As clsCode, _
                         ByVal d_from_file_name As String, _
                         ByVal d_from_file_title As String)
' ----------------------------------------------------------------------------
' Left is the current code, Right is the changed code
' ----------------------------------------------------------------------------
    Const PROC As String = "DiffFromDsply"
    
    On Error GoTo eh
    Dim flThis As File
    Dim flFrom As File
    
    Set flThis = AsFile("CurrentPublic")
    Set flFrom = d_from.AsFile("PendingRelease")
    
    DsplyFilesDiff d_file_left_full_name:=flThis.Path _
                 , d_file_right_full_name:=flFrom.Path _
                 , d_file_left_title:="The current public code (in Common-Components folder)" _
                 , d_file_right_title:="The changed code pending release"
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCode" & "." & sProc
End Function

Private Function FileFromArray(ByVal a_v As Variant, _
                               ByVal a_file_name As String, _
                      Optional ByVal f_ignore_empty As Boolean = True) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim ts          As TextStream
    Dim sFileName   As String
    
    With New FileSystemObject
        sFileName = .GetSpecialFolder(2) & "\" & a_file_name & ".txt"
        Set ts = .CreateTextFile(sFileName, True)
        If ArrayIsAllocated(a_v) Then
            ts.Write Join(a_v, vbCrLf)
        End If
        ts.Close
        Set FileFromArray = .GetFile(sFileName)
   End With

End Function

Private Sub Class_Initialize()
    bSourceIsNone = True
End Sub

Private Sub Class_Terminate()
    If Not flCode Is Nothing Then
        If fso.FileExists(flCode) Then fso.DeleteFile (flCode.Path)
    End If
End Sub

Public Function DsplyFilesDiff(ByVal d_file_left_full_name As String, _
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
        mWinMergeIni.Setup mWinMergeIni.WinMergeIniFullName
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


