VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit ' Under c o n s t r u c t i o n!
' ---------------------------------------------------------------------------
' Class Module clsProc: Provides methods and properties for the code of/in a
' ===================== Procedure (Sub, Function, Property) whereby the
' source may be a VB-Projects CodeModule or an Export-File. In any case the
' class deals with the/a mere procedure only.
'
' Attention: This class module treats a Properties Get/Let/Set procedures as
' ---------- one procedure which in fact embodies two procedures.
'
'
' Public Methods:
' ---------------
' CodeAsFile        Returns the mere code of a procedure (Sub, Function,
'               Property) as file.
' DiffersFrom   Returns TRUE when the code differs from the code of another
'               source.
' DsplyDiffs    Displays the difference of the code of two sources by means
'               of WinMerge.
' Meets         Not DiffersFrom
' ReplaceWith   Replaces the code in the provided CodeModule with the code
'               provided through another instance of this Class Module.
'
' Public Properties:
' ------------------
' AsString    r Provides the code as a single string with the code lines
'               delimited by a vbCrLf
' ProcId      r Name of the procedure within a component in the form
'               <comp-name>.<proc-name>.
' Source      w The source from which the procedure's code is to be
'               extracted, either a component's Export-File (as object or
'               full name) or a VB-Project's  CodeModule.
'
' W. Rauschenberger, Berlin May 2024
' ---------------------------------------------------------------------------
Private Const ERR_NO_SOURCE = "Obtaining the code as array or string when the source 'is none' is useless!" & vbLf & _
                              "The Property 'IsNone' may be used to avoid this. The Property is TRUE when " & _
                              "there are no known code lines for whichever reason, no source ever provided, " & _
                              "the source does not contain any code lines, or just empty code lines)."
Private vSource                         As Variant
Private sProcId                         As String
Private bSourceIsNone                   As Boolean
Private dctTempFiles                    As New Dictionary
Private sCode                           As String
Private sCompName                       As String
Private sProcName                       As String
Private sScope                          As String
Private aCode                           As Variant
Private sKindOfProc                     As String

Private Sub Class_Terminate()
    FileTempRemove
End Sub

Public Property Get Code() As String:               Code = sCode:               End Property

Public Property Let Code(ByVal s As String)
    If sCode = vbNullString Then
        sCode = s
    Else
        sCode = sCode & vbCrLf & s
    End If
    
End Property

Public Property Get CompName() As String:           CompName = sCompName:       End Property

Public Property Let CompName(ByVal s As String):    sCompName = s:              End Property

Public Property Get IsNone() As Boolean:            IsNone = bSourceIsNone: End Property

Public Property Get KindOfProc() As String:         KindOfProc = sKindOfProc:   End Property

Public Property Let KindOfProc(ByVal s As String):  sKindOfProc = s:            End Property

Public Property Get ProcId() As String:             ProcId = sProcId:           End Property

Public Property Get ProcName() As String:           ProcName = sProcName:       End Property

Public Property Let ProcName(ByVal s As String):    sProcName = s:              End Property

    
Public Property Get Scope() As String:              Scope = sScope:             End Property

Public Property Let Scope(ByVal s As String):       sScope = s:                 End Property

Public Property Let Source(ByVal f_v As Variant)
' ----------------------------------------------------------------------------
' Depending on the provided argument's (f_v) type the procedures code lines
' are saved as an array lines.
' Types acepted are: - String, considered a file's full name
'                    - File, considered a VBComponent's Export-File
'                    - VBComponent
'                    - CodeModule
' Note: The module is initialized with IsNone = TRUE and this is only set to
'       FALSE when a Source had been provided which contains at least one non
'       non empty code line.
' ----------------------------------------------------------------------------
    Const PROC = "Source-Let"
    
    Dim fl      As File
    Dim vbcm    As CodeModule
    
    If sProcName = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No procedure name has yet been provided - which is obligatory before a source " & _
                                            "may be provided from which the procedure's code lines are to be obtained!"
                                            
    Select Case TypeName(f_v)
        Case "String"
            If f_v <> vbNullString Then
                With FSo
                    If .FileExists(f_v) And sProcName <> vbNullString Then
                        vSource = f_v
                        Set fl = .GetFile(f_v)
                        sProcId = .GetBaseName(fl)
                        sCode = CodeLinesFromExpFile(fl, sProcName)
                    End If
                End With
            End If
        Case "VBComponent"
            Set vSource = f_v
            Set vbcm = f_v.CodeModule
            If vbcm.CountOfLines > 0 _
            Then sCode = CodeLinesFromCodeModule(vbcm, sProcName)
        Case "CodeModule"
            Set vSource = f_v
            Set vbcm = f_v
            sProcId = vbcm.Parent.Name
            If vbcm.CountOfLines > 0 _
            Then sCode = CodeLinesFromCodeModule(vbcm, sProcName)
        Case "File"
            Set vSource = f_v
            Set fl = f_v
            sProcId = FSo.GetBaseName(fl.Path)
            sCode = CodeLinesFromExpFile(f_v, sProcName)
    End Select
    
End Property

Public Function CodeAsFile(ByVal a_name As String) As File
' ----------------------------------------------------------------------------
' Returns the procedure's code lines (derived from a source) as temporary file
' with a given name (a_name).
' ----------------------------------------------------------------------------
    Dim sTempFile As String
    
    sTempFile = TempFile(a_name)
    With New FileSystemObject
        If .FileExists(sTempFile) _
        Then .DeleteFile sTempFile
    End With
    StringAsFile sCode, sTempFile
    
End Function

Private Function CodeLinesFromCodeModule(ByVal c_vbcm As CodeModule, _
                                         ByVal c_proc As String) As String
' ----------------------------------------------------------------------------
' Returns a string with all the code lines in the CodeModule (c_vbcm) whereby
' consecutive Property code lines are returned as one string.
' ----------------------------------------------------------------------------
    Dim lFirst  As Long
    Dim lLines  As Long
    
    If IsLocatedInCodeModule(c_vbcm, c_proc, lFirst, lLines) Then
        CodeLinesFromCodeModule = c_vbcm.Lines(lFirst, lLines)
    End If
    
End Function

Private Function CodeLinesFromExpFile(ByVal p_fl As File, _
                                      ByVal p_proc_name As String) As String
    Dim wbk As Workbook
    
    CodeLinesFromExpFile = CodeLinesFromCodeModule(TempCodeMod(p_fl.Path, wbk), p_proc_name)
    wbk.Close False
    
End Function

Public Function DiffersFrom(ByVal d_from_proc As clsProc) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when this class instance's code (sCode) differs from those of
' another's instance (d_from_proc).
' ----------------------------------------------------------------------------
    Const PROC = "DiffersFrom"
    
    On Error GoTo eh
    
    Select Case True
        Case Not IsNone And Not d_from_proc.IsNone
            DiffersFrom = StrComp(sCode _
                                , d_from_proc.Code _
                                , vbTextCompare) <> 0
        Case Not Me.IsNone And d_from_proc.IsNone:   DiffersFrom = True
        Case Me.IsNone And Not d_from_proc.IsNone:   DiffersFrom = True
        Case Me.IsNone And d_from_proc.IsNone:       Err.Raise AppErr(1), ErrSrc(PROC), ERR_NO_SOURCE
    End Select
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub DiffersFromDiffs(ByVal d_from_proc As clsProc)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "DiffersFromDiffs"
    
    Dim i       As Long
    Dim arr     As Variant
    Dim arrFrom As Variant
    
    Debug.Print ErrSrc(PROC) & ": " & "changed code: " & sCompName & "." & sProcName
    arr = Split(sCode, vbCrLf)
    arrFrom = Split(d_from_proc.Code, vbCrLf)
    For i = 1 To Min(UBound(arr), UBound(arrFrom))
        If StrComp(arr(i), arrFrom(i), vbTextCompare) <> 0 Then
            Debug.Print ErrSrc(PROC) & ": " & "                             The first difference has been detected in line " & i & ":"
            Debug.Print ErrSrc(PROC) & ": " & "                             Line " & i & " """ & arr(i) & """"
            Debug.Print ErrSrc(PROC) & ": " & "                             Line " & i & " """ & arrFrom(i) & """"
            Exit For
        End If
    Next i

End Sub

Public Sub DsplyDiffs(ByVal d_this_file_name As String, _
                      ByVal d_this_file_title As String, _
                      ByVal d_versus_proc As clsProc, _
                      ByVal d_versus_file_name As String, _
                      ByVal d_versus_file_title As String)
' ----------------------------------------------------------------------------
' Left is the procedure's current public code, Right is the procedure's code
' in the serviced Workbook's component.
' ----------------------------------------------------------------------------
    Const PROC As String = "DsplyDiffs"
    
    On Error GoTo eh
    Dim sFileThis   As String
    Dim sFileFrom   As String
    
    sFileThis = TempFile(d_this_file_name)
    sFileFrom = TempFile(d_versus_file_name)
    
    StringAsFile d_versus_proc.Code, sFileFrom
    StringAsFile sCode, sFileThis
    
    DsplyFilesDiff d_file_left_full_name:=sFileThis _
                 , d_file_right_full_name:=sFileFrom _
                 , d_file_left_title:=d_this_file_title _
                 , d_file_right_title:=d_versus_file_title
    
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
        
    If Not FSo.FileExists(d_file_left_full_name) _
    Then Err.Raise Number:=AppErr(2) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & d_file_left_full_name & """ does not exist!"
    
    If Not FSo.FileExists(d_file_right_full_name) _
    Then Err.Raise Number:=AppErr(3) _
                 , Source:=ErrSrc(PROC) _
                 , Description:="The file """ & d_file_right_full_name & """ does not exist!"
        
    '~~ Guarantee an ini-File which meets CompMan's specifics
    With FSo
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
    ErrSrc = "clsProc" & "." & sProc
End Function

Private Sub FileTempRemove()
' ------------------------------------------
' Remives all temporary files created.
' ------------------------------------------
    Dim v As Variant
    
    With New FileSystemObject
        For Each v In dctTempFiles
            If .FileExists(v) _
            Then .DeleteFile v
        Next v
    End With
    
End Sub

Public Function IsCopiedProc() As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Code of this class instance indicates a copied proc.
' ----------------------------------------------------------------------------
    Dim i As Long
    
    aCode = Split(sCode, vbCrLf)
    For i = LBound(aCode) To UBound(aCode)
        If InStr(aCode(i), sProcName & "(") <> 0 Then Exit For
        If aCode(i) Like PROC_COPY_INDICATOR Then
            IsCopiedProc = True
            Exit For
        End If
    Next i
    
End Function

Public Function IsLocatedInCodeModule(ByVal i_vbcm As CodeModule, _
                                      ByVal i_proc_name As String, _
                                      ByRef i_first_line As Long, _
                                      ByRef i_no_of_lines As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE, the first line (i_first_line), and the number of lines
' (i_no_of_lines) of a located procedure (i_proc_name).
' Note: When the kind of the procedure is a Property the number of lines cover
'       all consequtive code lines, i.e. Get, Let, Set.
' ----------------------------------------------------------------------------
    Dim i       As Long
    Dim lKind   As vbext_ProcKind
    Dim sProc   As String
    Dim lNext   As Long
    
    sProc = i_proc_name
    With i_vbcm
        i = .CountOfDeclarationLines + 1
        Do While i < .CountOfLines
            sProc = .ProcOfLine(i, lKind)
            If sProc = i_proc_name Then
                IsLocatedInCodeModule = True
                i_first_line = i
                i_no_of_lines = .ProcCountLines(sProc, lKind)
                If lKind = vbext_pk_Proc Then Exit Do
                
                lNext = i + .ProcCountLines(i_proc_name, lKind)
                sProc = .ProcOfLine(lNext, lKind)
                If sProc = i_proc_name And lKind <> vbext_pk_Proc Then
                    '~~ When the next procedure has the same name and is a kind of Property,
                    '~~ the number of codelines includes this Propertie's code lines
                    i_no_of_lines = i_no_of_lines + .ProcCountLines(sProc, lKind)
                    Exit Do
                End If
            Else
                i = i + .ProcCountLines(sProc, lKind)
            End If
        Loop
    End With
    
End Function

Public Sub ReplaceWith(ByVal r_target_vbcm As CodeModule, _
                       ByVal r_proc_name As String, _
                       ByVal r_source_code As clsProc)
' ----------------------------------------------------------------------------
' Replaces the code lines - if any - in the provided CodeModule (r_this_vbcm)
' with the code of the provided source (r_source_code).
' ----------------------------------------------------------------------------
    Dim lLine   As Long
    Dim lLines As Long
    
    If IsLocatedInCodeModule(r_target_vbcm, r_proc_name, lLine, lLines) Then
        With r_target_vbcm
            If .CountOfLines > 0 Then .DeleteLines lLine, lLines
            .InsertLines lLine, r_source_code.Code
        End With
    End If
    
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
    Set StringAsFile = FSo.GetFile(s_file)
    
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

Private Function TempCodeMod(ByVal t_file As String, _
                             ByRef t_wbk As Workbook) As CodeModule

    Dim vbc     As VBComponent
    
    Set t_wbk = Application.Workbooks.Add()
    With t_wbk
        .Windows(1).Visible = False
        With .VBProject
            .VBComponents.Import t_file
            Set vbc = .VBComponents(FSo.GetBaseName(t_file))
        End With
        Set TempCodeMod = vbc.CodeModule
    End With
    Set vbc = Nothing
    
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
        sTemp = Replace(FSo.GetTempName, ".tmp", f_extension)
    Else
        sTemp = f_name & f_extension
    End If
    If f_path = vbNullString Then f_path = FSo.GetSpecialFolder(2)
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp

End Function

