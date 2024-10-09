VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrivProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Class Module clsPrivProf: Private Profile Files services. The componet works
' ========================= autonomous and does not require any other
' components. Naming comvention: PPfile.... operates on the Private Profile
' file directly, PPtwin... operates on/with a Private Profile file's twin
' maintained in the form of a Directories/Collections structure.
'
' Public properties: (* with the exception of these properties all properties
'                       have arguments:
' ------------------
' FileName *)     r/w Specifies/returns the Private Profile file full name
'                     used as the default when the corresponding argument
'                     is omitted.
' FileFooter      r/w Reads/writes a file footer
' FileHeader      r/w Reads/writes a file footer
' Section *)      r/w Specifies/returns the current specified Section, used
'                     as default when the corresponding argument is omitted.
' SectionComment  r/w Reads(writes) a section's comments
' Value           r/w Reads/writes a value from/to a Private Profile File,
'                     optionally reads/writes a value comment
' ValueComment    r/w Reads(writes) a section's comments
'
' Public methods:
' ---------------
' Exists          Returns TRUE when a given value name exists in a given
'                 section exists in a given Private Profile File. The
'                 method requires a Private Profile file's full name,
'                 section and value name are optional.
' HskpngNames     Service supporting the fact that a value-name may change
'                 over time. The argument is a list of list of strings
'                 (ParamArray), where each string specifies a value-name
'                 history in the form:
'                 [[sect:][sect:]][current-name]<prev-name<prev-name or
'                 [[sect:][sect:]]prev-name>prev-name>[current-name]
'                 with the following syntax rules:
'                 1. > or < indicate the history sequence pointing from
'                    the predecessor to the new name, whereby the history
'                    may grow when a changed name is changed again.
'                 2. One or more sections are optional. When omitted, the
'                    names history applies to all sections any history
'                    name is used.
'                 3. The current name is optional. When omitted, the life of
'                    the history name has ended. I. e. the name is no longer
'                    used at all and thus will be removed in all or the
'                    specified sections only (rule 2).
' NamesRemove     Removes provided value names, in a given Private
'                 Properties File, when provided in a specific section,
'                 else in all sections.
' SectionNames    Returns all section names in a Private Profile file as
'                 Dictionary and optionally as Collection with all
'                 Section names in ascending order.
' Sections        Returns a Dictionary with all sections in a Private
'                 Profile file with the section name as the key.
'                 The item is a Collection where the first item is a
'                 Collection of the section comments and the second is a
'                 Dictionary of values.
' SectionNames    Returns all section names in a Private Profile file as
'                 Dictionary and optionally as Collection with all Section
'                 names in ascending order.
' SectionRemove   Removes one or more specified sections in a Private Profile
'                 file whereby several sections may be provided as a comma
'                 separated string. When all "user" sections had been removed
'                 the Private Profile file is also removed *).
' ValueNames      Returns a Dictionary with all value-names as key, whereby
'                 the item is a Collection with two items: the first is a
'                 Collection of the value comment lines, the second is the
'                 value. The value names are in ascending order.
' ValueRename     Function replaces an old value name with a new one
'                 either in a specific section or in all sections when no
'                 specific section is provided. Optionally not reorgs the
'                 file, returns True when at least one name has been
'                 replaced.
' ValueRemove     Removes a specified value from an (optionally) specified
'                 section in an (optionally) specified Private Profile file.
'                 When the section is omitted, the value is removed from all
'                 sections. When the file name is omitted it defaults to the
'                 file specified by the FileName property.
'
' *) The Private Profile file may contain a "names housekeeping section"
'    which is not regarded a "user" but a "system" section. I.e. when it is
'    the only section left, the file will be removed as well. This coinsides
'    with the fact that the file is written when the very first value is
'    written.
'
' Requires:
' ---------
' Reference to Microsoft Scripting Runtine
' Reference to Microsoft VBScript Regular Expressions 5.5
'
' Uses:
' -----
' No other components. Will use optionally mErH, fMsg/mMsg when installed
' and activated (Cond. Comp. Args. `mErH = 1 : mMsg = 1`).
'
' W. Rauschenberger, Berlin Oct 2024
' See also https://github.com/warbe-maker/VBA-Private-Profile.
' ----------------------------------------------------------------------------
Public FSo                      As FileSystemObject

Private Const WIN_NORMAL = 1         'Open Normal
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const NAMES_HSKPNG_DONE_DELIM   As String = ";"
Private Const NAMES_HSKPNG_SECT_NAME    As String = "@NamesHousekeeping"
Private Const NAMES_HSKPNG_VALUE_NAME   As String = "DoneNamesHskpng"
Private Const VALUE_TRUE                As String = "TRUE"
Private Const VALUE_FALSE               As String = "FALSE"

#If mMsg = 0 Then
    ' ------------------------------------------------------------------------
    ' The 'minimum error handling' approach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed 1) which must
    ' be indicated by the Conditional Compile Argument mMsg = 1
    '
    ' 1) See https://github.com/warbe-maker/Common-VBA-Message-Service for
    '    how to install an use.
    ' ------------------------------------------------------------------------
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Private Declare PtrSafe Function apiShellExecute _
                Lib "shell32.dll" Alias "ShellExecuteA" _
               (ByVal hWnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Private bBooleanAsTrueFalseString   As Boolean      ' option, when TRUE a Boolean value is saved as a TRUE or FALSE string
Private bSectionsSeparated          As Boolean      ' section separator line option, defaults to TRUE
Private dctPPtwinsAsString          As Dictionary   ' file as string
Private Reg                         As RegExp
Private sFileFooter                 As String
Private sFileFooterDelimiterChar    As String       ' the footer delimiter character (must be different from the above)
Private sFileHeader                 As String
Private sFileHeaderDelimiterChar    As String       ' the header delimiter character
Private sFileName                   As String       ' the current default Private Profile file's full name
Private sSection                    As String       ' the current section of the above Private Profile file
' -----------------------------------------------------------------------------------------------------------------------------------------
' Private elements of Private Profile file twins structure used throughout this component for twin actions (indentaions = structure)
Private dctPPtwins                                     As Dictionary ' All Private Profile file twins in a class instance (key = file-name)
    Private cllPPtwin                                  As Collection ' A Private Profile file's twin
        Private cllPPtwinHeader                        As Collection ' A files's header (may be 0)
        Private dctPPtwinSections                      As Dictionary ' All sections (key = section)
            Private cllPPtwinSection                   As Collection ' A section
                Private cllPPtwinSectionComments       As Collection ' A section's comments (may be 0, 1 when separator is True)
                Private dctPPtwinValues                As Dictionary ' A section's values (key = value-name)
                    Private cllPPtwinValue             As Collection
                        Private cllPPtwinValueComments As Collection ' A values comments (may be 0)
                        Private vPPtwinValue           As Variant    ' The value
        Private cllPPtwinFooter                        As Collection ' The file's footer comments (may be 0)
' -----------------------------------------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
    
    Set Reg = New RegExp
    Set FSo = New FileSystemObject
    Set dctPPtwins = New Dictionary
    Set dctPPtwinsAsString = New Dictionary
    
    If sFileName = vbNullString Then
        '~~ Initialize the default Private Profile File name
        Select Case Application.Name
            Case "Microsoft Excel": sFileName = ThisWorkbook.Path & "\" & FSo.GetBaseName(ThisWorkbook.Name) & ".dat"
            Case Else:              Err.Raise AppErr(1), ErrSrc(PROC), "Initializing the FileName failed!" & vbLf & _
                                                                       "The current Application.Name is not Excel " & _
                                                                       "and any other application is yet not supported!"
        End Select
    End If
    sFileHeaderDelimiterChar = "="
    sFileFooterDelimiterChar = "-"
    bBooleanAsTrueFalseString = True
    bSectionsSeparated = True
    
End Sub

Private Sub Class_Terminate()
    Set FSo = Nothing
    Set dctPPtwins = Nothing
End Sub

Private Property Get Arry(Optional ByRef c_arr As Variant, _
                          Optional ByVal c_index As Long = -1) As Variant
' ----------------------------------------------------------------------------
' Universal array read procedure. Returns Null when a given array (c_arr) is
' not allocated or a provided index is beyond/outside current dimensions.
' ----------------------------------------------------------------------------
    Dim i As Long
    
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
' Universal array add/update procedure, avoiding any prior checks whether
' allocated, empty not yet existing, etc.
' - Adds an item (c_var) to an array (c_arr) when no index is provided or when
'   the index lies beyond UBound
' - When an index is provided, the item is inserted/updated at the given
'   index - even when the array yet doesn't exist or is not yet allocated.
' ----------------------------------------------------------------------------
    Const PROC = "Arry-Let"
    
    Dim bIsAllocated As Boolean
    Dim s            As String
    
    If IsArray(c_arr) Then
        On Error GoTo -1
        On Error Resume Next
        bIsAllocated = UBound(c_arr) >= LBound(c_arr)
        On Error GoTo eh
    ElseIf VarType(c_arr) <> 0 Then
        Err.Raise AppErr(1), ErrSrc(PROC), "Not a Variant type!"
    End If
    
    If bIsAllocated = True Then
        '~~ The array has at least one item
        If c_index = -99 Then
            '~~ When for an allocated array no index is provided, the item is added
            ReDim Preserve c_arr(UBound(c_arr) + 1)
            c_arr(UBound(c_arr)) = c_var
        ElseIf c_index >= 0 And c_index <= UBound(c_arr) Then
            '~~ Replace an existing item
            c_arr(c_index) = c_var
        ElseIf c_index > UBound(c_arr) Then
            '~~ New item beyond current UBound
            ReDim Preserve c_arr(c_index)
            c_arr(c_index) = c_var
        ElseIf c_index < LBound(c_arr) Then
            Err.Raise AppErr(2), ErrSrc(PROC), "Index is less than LBound of array!"
        End If
        
    ElseIf bIsAllocated = False Then
        '~~ The array does yet not exist
        If c_index = -99 Then
            '~~ When no index is provided the item is the first of a new array
            c_arr = Array(c_var)
        ElseIf c_index >= 0 Then
            ReDim c_arr(c_index)
            c_arr(c_index) = c_var
        Else
            Err.Raise AppErr(3), ErrSrc(PROC), "the provided index is less than 0!"
        End If
    End If
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let BooleanAsTrueFalseString(ByVal b As Boolean):   bBooleanAsTrueFalseString = b:  End Property

Public Property Get FileFooter(Optional ByVal name_file As String = vbNullString) As String
' ------------------------------------------------------------------------------
' Returns the header of a file as a strings delimited by vbCrLf, ",", ", ", ";",
' or "; ".
' ------------------------------------------------------------------------------
    Const PROC = "SectionComment-Get"
    
    Dim cll As New Collection
    Dim v   As Variant
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    
    For Each v In PPtwinFileFooter(name_file)
        If Not IsDelimiterFileFooter(v) Then
            cll.Add Replace(v, "; ", vbNullString, , 1)
        End If
    Next v
    
    FileFooter = CollectionAsString(cll)
    Set cll = Nothing

End Property

Public Property Let FileFooter(Optional ByVal name_file As String = vbNullString, _
                                        ByVal file_footer As String)
' ------------------------------------------------------------------------------
' Write a header for a file (name_fiel) provided as a strings delimited by
' vbCrLf, ",", ", ", ";", or "; ".
' ------------------------------------------------------------------------------
    Const PROC = "SectionComment-Get"
    
    Dim sFooter As String
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not FileIsActive(name_file) Then GoTo xt

    '~~ Make sure the provided footer lines are delimeted by vbCrLf
    sFooter = Replace(Replace(file_footer, vbLf, vbCrLf), vbCr & vbCrLf, vbCrLf)
    
    If sFileFooter <> sFooter Then
        PPtwinFileFooter(name_file) = FileFooterDelimiterAdd(StringAsCollection(sFooter))
        PPmanage name_file ' in case no value is changed or new
        sFileFooter = sFooter ' will be set when the file has become active
    End If
    
xt: Exit Property
End Property

Public Property Get FileHeader(Optional ByVal name_file As String = vbNullString) As String
' ------------------------------------------------------------------------------
' Returns the header of a file as a strings delimited by vbCrLf, ",", ", ", ";",
' or "; " without a header delimiter.
' ------------------------------------------------------------------------------
    Const PROC = "FileHeader-Get"
    
    On Error GoTo eh
    Dim cll As New Collection
    Dim v   As Variant
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    
    For Each v In PPtwinFileHeader(name_file)
        If Not IsDelimiterFileHeader(v) Then
            cll.Add Replace(v, "; ", vbNullString, , 1)
        End If
    Next v
    
    FileHeader = CollectionAsString(cll)
    Set cll = Nothing
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Function FileIsActive(ByVal name_file As String) As Boolean
' ------------------------------------------------------------------------------
' The function returns TRUE when the file (name_file) exists and has at least
' one user section - with at least one value.
' ------------------------------------------------------------------------------
    Dim dctSects As Dictionary
    
    If Not FSo.FileExists(name_file) Then Exit Function
    PPmanage name_file
    Set dctSects = SectionNames(name_file)
    Select Case True
        Case dctSects.Count = 0
        Case dctSects.Count = 1 And dctSects.Keys(0) = NAMES_HSKPNG_SECT_NAME
            ValueRemove NAMES_HSKPNG_VALUE_NAME, NAMES_HSKPNG_SECT_NAME, name_file ' without any user section history is obsolete
        Case Else
            FileIsActive = True
    End Select
    Set dctSects = Nothing
    
End Function

Public Property Let FileHeader(Optional ByVal name_file As String, _
                                        ByVal file_header As String)
' ------------------------------------------------------------------------------
' Write a header for a file (name_fiel) provided as a strings delimited by
' vbCrLf, ",", ", ", ";", or "; " by adding a header delimiter when not alrteady
' provided.
' ------------------------------------------------------------------------------
    Const PROC = "FileHeader-Let"
    
    Dim sHeader As String
    Dim i       As Long
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not FileIsActive(name_file) Then GoTo xt
        
    '~~ Make sure the provided header lines are delimeted by vbCrLf
    sHeader = Replace(Replace(file_header, vbLf, vbCrLf), vbCr & vbCrLf, vbCrLf)
    
    If sFileHeader <> sHeader Then
        PPtwinFileHeader(name_file) = FileHeaderDelimiterAdd(StringAsCollection(sHeader))
        PPmanage name_file ' in case no value is changed or new
        sFileHeader = sHeader ' will be set when the file has become active
    End If
    
xt: Exit Property
End Property

Public Property Get FileName() As String:               FileName = sFileName:       End Property

Public Property Let FileName(ByVal f_file As String)
' ----------------------------------------------------------------------------
' Note: Only a FileName explicitely provided as Property becomes the default
'       file name for an omitted file argument.
'       Any file name provided as argument along with a property or method
'       call is only valid for the call.
' ----------------------------------------------------------------------------
    Const PROC = "FileName-Let"
    
    On Error GoTo eh
    Dim sInvalidLine    As String
    Dim arr             As Variant
    
    If Not IsEmptyFile(f_file) _
    And Not IsValidFileContent(f_file, arr, sInvalidLine) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidFileContent(f_file, sInvalidLine)
    
    sFileName = f_file
    PPmanage f_file
    sFileHeader = FileHeader(f_file) ' current header
    sFileFooter = FileFooter(f_file) ' current footer

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get Footer(Optional ByVal name_file As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' When the file has no footer an empty Collection is returned. A footer is
' at least one comment line with a footer delimiter lin above.
' ----------------------------------------------------------------------------
    
    If PPtwinExists(name_file) Then ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        Set Footer = cllPPtwinFooter
    End If
    
End Property

Public Property Let Footer(Optional ByVal name_file As String = vbNullString, _
                                    ByVal file_footer As Variant)
' ----------------------------------------------------------------------------
' Replaces an old footer in file (name_file) - if any - with the new footer
' (file_footer) - which may be a series of strings delimited by vbCrLf, vbLf,
' "|" or "||".
' ----------------------------------------------------------------------------
    Const PROC = "Footer-Let"
        
    On Error GoTo eh
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If PPtwinExists(name_file) Then         ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        PPtwinFileFooter(name_file) = file_footer
        PPmanage name_file
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), "Writing a file footer into a yet not existing file is not supported! " & _
                                           "A footer may be written with any Value Property or explicitely once " & _
                                           "at least one value had been written."
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get NameHskpngSection() As String: NameHskpngSection = NAMES_HSKPNG_SECT_NAME: End Property

Private Property Get PPtwinFileFooter(ByVal p_file As String) As Collection
    
    Set PPtwinFileFooter = New Collection
    If PPtwinExists(p_file) Then    ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        Set PPtwinFileFooter = cllPPtwinFooter
    End If
    
End Property

Private Property Let PPtwinFileFooter(Optional ByVal p_file As String = vbNullString, _
                                           ByVal p_comments As Variant)
' ----------------------------------------------------------------------------
' Replaces or adds a file footer to a Private Profile file's (p_file) twin.
' Note: The file needs to be existing, i.e. it must have at least one section
'       with one value
' ----------------------------------------------------------------------------
        
    If PPtwinExists(p_file) Then                    ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        Set cllPPtwinFooter = PPtwinFileFooterPrepare(p_comments)
        PPtwinReplace p_file
    End If

End Property

Private Property Get PPtwinFileHeader(ByVal p_file As String) As Collection
    Set PPtwinFileHeader = New Collection
    If PPtwinExists(p_file) Then    ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        Set PPtwinFileHeader = cllPPtwinHeader
    End If
End Property

Private Property Let PPtwinFileHeader(Optional ByVal p_file As String = vbNullString, _
                                               ByVal p_comments As Collection)
' ----------------------------------------------------------------------------
' Replaces any existing file comments with new ones (p_comments),
' whereby the new comments (p_comments) may be provided as a series
' of strings delimited by vbCrLf, vbLf, "|" or "||", a Collection of
' strings, or a Dictionary with the comments a key.
' ----------------------------------------------------------------------------
        
    If PPtwinExists(p_file) Then                        ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        Set cllPPtwinHeader = FileHeaderDelimiterAdd(p_comments)
        PPtwinReplace p_file
        PPmanage p_file
    End If

End Property

Private Property Let PPtwinHskpngNamesDone(Optional ByVal h_file As String, _
                                                    ByVal h_hist As String)
' ----------------------------------------------------------------------------
' Keeps a record of a done name housekeeping (h_hist) in a Private Profile
' file (h_file).
' ----------------------------------------------------------------------------
    Dim lNew As Long
    
    If Not PPtwinHskpngNamesIsDone(h_file, h_hist) Then
        lNew = PPtwinHskpngNamesNames(h_file).Count + 1
        PPtwinValue(NAMES_HSKPNG_VALUE_NAME & "_" & lNew, NAMES_HSKPNG_SECT_NAME, h_file) = h_hist
    End If
    
End Property

Private Property Get PPtwinSectionComments(Optional ByVal p_sect As String, _
                                           Optional ByVal p_file As String) As Collection
' ----------------------------------------------------------------------------
' Returns an existing sections comments excluding the empty separator line
' when SectionSeparated = True.
' ----------------------------------------------------------------------------
    Set PPtwinSectionComments = New Collection
    
    If PPtwinSectionExists(p_sect, p_file) Then ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
        Set PPtwinSectionComments = cllPPtwinSectionComments
    End If
    
End Property

Private Property Let PPtwinSectionComments(Optional ByVal p_sect As String, _
                                           Optional ByVal p_file As String, _
                                                    ByVal p_comments As Collection)
' ----------------------------------------------------------------------------
' Writes comments to a section.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinSectionComments"
    
    On Error GoTo eh
    If PPtwinSectionExists(p_sect, p_file) Then ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
        Set cllPPtwinSectionComments = PPtwinSectionCommentsLet(p_comments)
        PPtwinReplaceSection p_sect
        PPtwinReplace p_file
        PPmanage p_file
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get PPtwinValue(Optional ByVal p_name As String, _
                                 Optional ByVal p_sect As String, _
                                 Optional ByVal p_file As String, _
                                 Optional ByVal p_default As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    If IsMissing(p_default) Then p_default = vbNullString
    
    If PPtwinValueExists(p_name, p_sect, p_file) Then
        PPtwinValue = PPvalue(vPPtwinValue, p_default)
    Else
        PPtwinValue = p_default
    End If
            
End Property

Private Property Let PPtwinValue(Optional ByVal p_name As String, _
                                 Optional ByVal p_sect As String, _
                                 Optional ByVal p_file As String, _
                                 Optional ByVal p_default As Variant, _
                                          ByVal p_valu As Variant)
' ----------------------------------------------------------------------------
' Modifies or adds a value (p_name=p_value) in a section (p_sect) in a
' Private Profile file's twin (p_file) whereby anything missing - section and
' or value is created.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinValue-Let"
    
    On Error GoTo eh
    If PPtwinExists(p_file) Then                    ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        If dctPPtwinSections.Exists(p_sect) Then
            PPtwinSectionSet p_sect                 ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
            If dctPPtwinValues.Exists(p_name) Then
                PPtwinValueSet p_name               ' cllPPtwinValue, cllPPtwinValueComments, vPPtwinValue
                vPPtwinValue = PPvalue(p_valu)
                PPtwinReplaceValue p_name           ' the/any comment remains unchanged
                PPtwinReplaceSection p_sect
                PPtwinReplace p_file
            Else
                '~~ New value
                If dctPPtwinValues Is Nothing Then Set dctPPtwinValues = New Dictionary
                dctPPtwinValues.Add p_name, PPtwinNewValue(PPvalue(p_valu))
                PPtwinReplaceSection p_sect
                PPtwinReplace p_file
                
            End If
        Else
            '~~ New section, new value
            If dctPPtwinSections Is Nothing Then Set dctPPtwinSections = New Dictionary
            dctPPtwinSections.Add p_sect, PPtwinNewSection(p_name, p_sect, p_valu)
            PPtwinReplace p_file
        End If
    Else
        ' New file, new section, new value
        dctPPtwins.Add p_file, PPtwinNew(p_name, p_sect, p_valu)
    End If
    
    If FileIsActive(p_file) Then
        '~~ When the file has become active for the very first time - i.e. when the first user value is written -
        '~~ a provided file header/footer is inserted.
        '~~ !! For an active file this is done immediately when the header/footer is provided. Only !!
        '~~ !! when the file at that point was not active the insertion is postponed to this point. !!
        If FileHeader(p_file) <> sFileHeader Then
            FileHeader(p_file) = sFileHeader
        End If
        If FileFooter(p_file) <> sFileFooter Then
            FileFooter(p_file) = sFileFooter
        End If
    End If

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get PPtwinValueComments(Optional ByVal p_name As String, _
                                         Optional ByVal p_sect As String, _
                                         Optional ByVal p_file As String) As Collection
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Set PPtwinValueComments = New Collection
    
    If PPtwinValueExists(p_name, p_sect, p_file) Then   ' cllPPtwinValue, cllPPtwinValueComments, vPPtwinValue
        Set PPtwinValueComments = cllPPtwinValueComments
    End If

End Property

Private Property Let PPtwinValueComments(Optional ByVal p_name As String, _
                                         Optional ByVal p_sect As String, _
                                         Optional ByVal p_file As String, _
                                                  ByVal p_comments As Collection)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    
    If PPtwinValueExists(p_name, p_sect, p_file) Then   ' cllPPtwinValue, cllPPtwinValueComments, vPPtwinValue
        Set cllPPtwinValueComments = PPtwinValueCommentsLet(p_comments)
        PPtwinReplaceValue p_name
        PPtwinReplaceSection p_sect
        PPtwinReplace p_file
        PPmanage p_file
    End If

End Property

Public Property Get Section() As String:        Section = sSection: End Property

Public Property Let Section(ByVal s As String)
' ----------------------------------------------------------------------------
' Note: A provided valid Section name becomes the default when subsequently
'       the corresponding argument is omitted.
' ----------------------------------------------------------------------------
    Const PROC = "Section-Let"
    
    On Error GoTo eh
    If Not IsValidNameSection(s) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidSectName(s)
    sSection = s
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get SectionComment(Optional ByVal name_section As String = vbNullString, _
                                   Optional ByVal name_file As String = vbNullString) As String
' ------------------------------------------------------------------------------
' Returns the comments of a section as a strings delimited by
' vbCrLf, ",", ", ", ";", "; ".
' ------------------------------------------------------------------------------
    Const PROC = "SectionComment-Get"
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidNameSection(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
                                     
    SectionComment = CommentsGet(PPtwinSectionComments(name_section, name_file))

End Property

Public Property Let SectionComment(Optional ByVal name_section As String = vbNullString, _
                                   Optional ByVal name_file As String = vbNullString, _
                                            ByVal value_comment As String)
' ------------------------------------------------------------------------------
' Returns the comments of a section as a vbCrLf delimited string.
' ------------------------------------------------------------------------------
    Const PROC = "SectionComment-Let"
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidNameSection(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
                                     
    PPtwinSectionComments(name_section, name_file) = StringAsCollection(value_comment)
    PPmanage name_file
    
End Property

Public Property Let SectionsSeparated(ByVal b As Boolean):          bSectionsSeparated = b:         End Property

Public Property Get Value(Optional ByVal name_value As String = vbNullString, _
                          Optional ByVal name_section As String = vbNullString, _
                          Optional ByVal name_file As String = vbNullString, _
                          Optional ByVal value_default As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns value name's (name_value) value from a section (name_section) in a
' Private Profile file (name_file) - plus its comment (it any) as a string of
' comment lines delimited with a "||".
' When no file (name_file) is provided it defaults to the file name specified
' by the property FileName, when no section (name_section) is provided, the
' section defaults to the section name provided by the property Section.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Value-Get"
           
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidNameSection(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    If Not IsValidNameValue(name_value) Then Err.Raise AppErr(3), ErrSrc(PROC), ErrInvalidValuName(name_value)
    
    Value = PPtwinValue(name_value, name_section, name_file, value_default)
    
End Property

Public Property Let Value(Optional ByVal name_value As String = vbNullString, _
                          Optional ByVal name_section As String = vbNullString, _
                          Optional ByVal name_file As String = vbNullString, _
                          Optional ByVal value_default As Variant, _
                                   ByVal v_value As Variant)
' ----------------------------------------------------------------------------
' Writes a value (v_value) under a given name (name_value) into a section
' (name_section) in a Private Profile file (name_file).
' When no file (name_file) is provided it defaults to the file name specified
' by the property FileName, when no section (name_section) is provided, the
' section defaults to the section name provided by the property Section.
' ----------------------------------------------------------------------------
    Const PROC = "Value-Let"
        
    On Error GoTo eh
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidNameSection(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    If Not IsValidNameValue(name_value) Then Err.Raise AppErr(3), ErrSrc(PROC), ErrInvalidValuName(name_value)
    
    PPtwinValue(name_value, name_section, name_file) = v_value
    PPmanage name_file ' re-setup the Private Profile file's twin
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get ValueComment(Optional ByVal name_value As String = vbNullString, _
                                 Optional ByVal name_section As String = vbNullString, _
                                 Optional ByVal name_file As String = vbNullString) As String
' ------------------------------------------------------------------------------
' Returns the comments of a value as a vbCrLf delimited string.
' ------------------------------------------------------------------------------
    Const PROC = "ValueComment-Get"
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidNameSection(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    If Not IsValidNameValue(name_value) Then Err.Raise AppErr(3), ErrSrc(PROC), ErrInvalidValuName(name_value)
                                     
    ValueComment = CommentsGet(PPtwinValueComments(name_value, name_section, name_file))

End Property

Public Property Let ValueComment(Optional ByVal name_value As String = vbNullString, _
                                 Optional ByVal name_section As String = vbNullString, _
                                 Optional ByVal name_file As String = vbNullString, _
                                          ByVal value_comment As String)
' ------------------------------------------------------------------------------
' Writes comments for a value provided as strings delimited by
' vbCrLf, ",", ", ", ";", "; ".
' ------------------------------------------------------------------------------
    Const PROC = "ValueComment-Let"
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidNameSection(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    If Not IsValidNameValue(name_value) Then Err.Raise AppErr(3), ErrSrc(PROC), ErrInvalidValuName(name_value)
                                     
    PPtwinValueComments(name_value, name_section, name_file) = StringAsCollection(value_comment)
    PPmanage name_file
    
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

' Origin: Common Component "mVarTrans" =======================================
Private Function ArrayAsCollection(ByVal a_array As Variant) As Collection
' ----------------------------------------------------------------------------
' Return an array's (a_array) items as Collection.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim v   As Variant
    
    With cll
        For Each v In a_array
            .Add v
        Next v
    End With
    Set ArrayAsCollection = cll
    Set cll = Nothing
    
End Function

' Origin: Common Component "mVarTrans" =======================================
Private Function ArrayAsDictionary(ByVal a_array As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the array's items are returned as Directory keys, the
'            items will be unified. I e. each item will exist only once. To
'            make this restriction productive, the number of occurrences of
'            each item is returned as item.'
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    Dim v As Variant
    
    With dct
        For Each v In a_array
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .Item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set ArrayAsDictionary = dct
    Set dct = Nothing
    
End Function

' Origin: Common Component "mVarTrans" =======================================
Private Function ArrayAsFile(ByVal a_array As Variant, _
                   Optional ByRef a_file As Variant = vbNullString, _
                   Optional ByVal a_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes all items of an array (a_arry) to a file (a_file) which might be a
' file object, a file's full name. When no file (a_file) is provided a
' temporary file is returned, else the provided file (a_file) as object.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
      
    If Not ArrayIsAllocated(a_array) Then Exit Function
    
    Select Case True
        Case a_file = vbNullString:     a_file = TempFile
        Case TypeName(a_file) = "File": a_file = a_file.Path
    End Select
    
    If a_file_append _
    Then Open a_file For Append As #1 _
    Else Open a_file For Output As #1
    Print #1, Join(a_array, vbCrLf)
    Close #1
    Set ArrayAsFile = FSo.GetFile(a_file)
    
End Function

' Origin: Common Component "mVarTrans" =======================================
Private Function ArrayAsString(ByVal a_array As Variant, _
                     Optional ByVal a_delim As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns an array (a_array) as string with the items delimited (a_delim).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    ArrayAsString = Join(a_array, a_delim)
End Function

Private Function ArrayIsAllocated(ByVal Arry As Variant) As Boolean
    
    On Error Resume Next
    ArrayIsAllocated = UBound(Arry) >= LBound(Arry)
    On Error GoTo -1
    
End Function

Private Sub BoC(ByVal b_id As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Bnd-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.BoC b_id, b_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.BoC b_id, b_args
#End If
End Sub

Private Sub BoP(ByVal b_proc As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP b_proc, b_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP b_proc, b_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

' Origin: Common Component "mVarTrans" =======================================
Private Function CollectionAsArray(ByVal c_coll As Collection) As Variant
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr     As Variant
    Dim sName   As String
    Dim v       As Variant
    
    For Each v In c_coll
        If IsObject(v, sName) _
        Then Arry(arr) = sName _
        Else Arry(arr) = v
    Next v
    CollectionAsArray = arr
    
End Function

' Origin: Common Component "mVarTrans" =======================================
Private Function CollectionAsDictionary(ByVal c_coll As Collection) As Dictionary
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as Dictionary keys.
' Attention: Because the collection's items are returned as Directory keys,
'            the items will be unified. I e. each item will exist only once.
'            To make this restriction productive, the number of occurrences of
'            each item is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    Dim v As Variant
    
    With dct
        For Each v In c_coll
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .Item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set CollectionAsDictionary = KeySort(dct)
    Set dct = Nothing

End Function

' Origin: Common Component "mVarTrans" =======================================
Private Function CollectionAsFile(ByVal c_coll As Collection, _
                         Optional ByRef c_file As String = vbNullString, _
                         Optional ByVal c_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Transfers the items of a Collection (c_coll) as records/lines to a
' file (c_file), optionally appended (c_append).
' ----------------------------------------------------------------------------

    If c_file = vbNullString Then c_file = TempFile
    StringAsFile CollectionAsString(c_coll, vbCrLf), c_file, c_append
    Set CollectionAsFile = FSo.GetFile(c_file)

End Function

Private Function CollectionAsString(ByVal c_coll As Collection, _
                           Optional ByVal c_delimiter As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as string with the items delimited
' by a vbCrLf. Itmes are converted into a string, if an item is an object its
' Name property is used (an error is raised when the object has no Name
' property.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sDelim  As String
    Dim sName   As String
    Dim v       As Variant
    
    If Not c_coll Is Nothing Then
        For Each v In c_coll
            If IsObject(v, sName) _
            Then s = s & sDelim & sName _
            Else s = s & sDelim & v
            sDelim = c_delimiter
        Next v
    End If
    CollectionAsString = s

End Function

Private Function Comment(ByVal s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s) as a propert Private Profile file comment line by
' considering that an empty line is a valid comment line.
' ----------------------------------------------------------------------------
    If Len(Trim(s)) = 0 Then
        Comment = Trim(s)   ' accepts an empty line as a valid quasi comment
    Else
        s = Replace(";" & s, ";;", ";")
        Comment = s         ' ensures, the line begins with a single ;
        If Left(Comment, 2) <> "; " Then
            Comment = "; " & Right(Comment, Len(Comment) - 1)
        End If
    End If
    
End Function

Private Function CommentsGet(ByVal c_comments As Collection) As String
' ------------------------------------------------------------------------------
' Returns a collection (c_comments) with a first empty item ignored and all
' the other items with the comment prefix "; " unstripped.
' ------------------------------------------------------------------------------

    Dim s       As String
    Dim sDelim  As String
    Dim v       As Variant
    
    For Each v In c_comments
        If v = vbNullString And s = vbNullString Then
            '~~ A first empty comment is ignored
        Else
            s = s & sDelim & Right(v, Len(v) - 2)
            sDelim = vbCrLf
        End If
    Next v
    CommentsGet = s
    
End Function

Private Function CommentsOfItem(ByVal h_arr As Variant, _
                              ByRef h_idx As Long) As Collection
    Dim cll As New Collection
    
    If h_idx = LBound(h_arr) Then GoTo xt
    h_idx = h_idx - 1
    Do While IsComment(h_arr(h_idx))
        cll.Add h_arr(h_idx)
        If h_idx = LBound(h_arr) Then Exit Do
        h_idx = h_idx - 1
    Loop

xt: Set CommentsOfItem = RevertOrder(cll)
    Set cll = Nothing
    
End Function

Public Function DictionaryAsArray(ByVal v_dict As Dictionary) As Variant
' ----------------------------------------------------------------------------
' Returns the keys of a Dictionary (v_dict) as Array.
' Precondition: All keys a convertable to a string.
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim v   As Variant
    
    For Each v In v_dict
        Arry(arr) = CStr(v)
    Next v
    DictionaryAsArray = arr
    
End Function

Private Function DictionaryAsString(ByVal v_dict As Dictionary, _
                           Optional ByVal v_delimiter As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns a Dictionary's (v_dict) keys as string, by default delimited by a
' "," (komma).
' Precondition: All keys are a - or are convertable to a - string.
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sDelim  As String
    Dim v       As Variant
    
    For Each v In v_dict
        s = s & sDelim & v
        sDelim = v_delimiter
    Next v
    DictionaryAsString = s

End Function

Private Function DictionaryInverted(ByVal d_dict As Dictionary) As Dictionary
    Dim dct As New Dictionary
    Dim v   As Variant
    
    For Each v In d_dict
        If Not dct.Exists(d_dict(v)) Then dct.Add d_dict(v), v
    Next v
    Set DictionaryInverted = dct
    Set dct = Nothing
    
End Function

Private Sub EoC(ByVal e_id As String, _
       Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End-of-Code' interface for the Common VBA Execution Trace Service.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mTrc = 1 Then         ' when mTrc is installed and active
    mTrc.EoC e_id, e_args
#ElseIf clsTrc = 1 Then   ' when clsTrc is installed and active
    Trc.EoC e_id, e_args
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP e_proc, e_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub

Private Function ErrInvalidFileContent(ByVal e_file As String, _
                                       ByVal e_line As String) As String
    ErrInvalidFileContent = "The file """ & e_file & """ has an invalid content ('" & e_line & """!"
End Function

Private Function ErrInvalidFileName(ByVal s As String) As String
    ErrInvalidFileName = "The provided file name (" & s & ") is not a valid Private Profile file!" & vbLf & _
                         "Valid is: - a not yet existing file" & vbLf & _
                         "          - an existing but empty file" & vbLf & _
                         "          - an existing file of which the content adheres to Private Profile file content rules." & vbLf & _
                         "            (see " & _
                         "has one been specified explicitely through the FileName property!"
End Function

Private Function ErrInvalidSectName(ByVal s As String) As String
    ErrInvalidSectName = "The provided section name (" & s & ") is invalid! " & _
                         "In a Private Profile files a section names may only " & _
                         "contain the characters A-Z, a-z, 0-9, _ (underscore), and @ (at) !"
End Function

Private Function ErrInvalidValuName(ByVal s As String) As String
    ErrInvalidValuName = "The provided value name (" & s & ") is invalid! " & _
                         "In a Private Profile file a value name may only " & _
                         "contain the characters A-Z, a-z, 0-9, and _ whereby the length is >= 1 <= 32 characters!"
End Function

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
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf mMsg = 1 Then
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
    If err_source = vbNullString Then err_source = Err.Source
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
            Else: ErrType = "VB Runtime Error "
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
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsPrivProf." & sProc
End Function

Public Function Exists(ByVal name_file As String, _
              Optional ByVal name_section As String = vbNullString, _
              Optional ByVal name_value As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the provided value name, section name, file name exists.
' When the provided (section (name_section
' ----------------------------------------------------------------------------
    Const PROC = "Exists"
    
    On Error GoTo eh
    Dim arr     As Variant
    Dim bActive As Boolean
    
    If Not IsValidFile(name_file, arr, bActive) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not bActive Then GoTo xt
        
    Select Case True
        Case name_section <> vbNullString And name_value = vbNullString:    Exists = PPtwinSectionExists(name_section, name_file)
        Case name_section <> vbNullString And name_value <> vbNullString:   Exists = PPtwinValueExists(name_value, name_section, name_file)
    End Select
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsArray(ByVal f_file As Variant, _
                    Optional ByVal f_empty_excluded As Boolean = False) As String()
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "FileAsArray"
    
    On Error GoTo eh
    Dim arr   As Variant
    Dim v     As Variant
    Dim s     As String
    Dim sSplit As String
    
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    If Not FSo.FileExists(f_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named """ & f_file & """ does not exist!"
    
    s = FileAsString(f_file, sSplit)
    If f_empty_excluded Then
        For Each v In Split(s, sSplit)
            If Trim$(v) <> vbNullString Then Arry(arr) = v
        Next v
    Else
        arr = Split(s, sSplit)
    End If
    FileAsArray = arr
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

' Origin: Common Component "mVarTrans" =======================================
Private Function FileAsCollection(ByVal f_file As Variant) As Collection
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as Collection.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsCollection"
    
    On Error GoTo eh
    Dim cll As New Collection
    Dim v   As Variant
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    If Not FSo.FileExists(f_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named """ & f_file & """ does not exist!"

    With cll
        For Each v In FileAsArray(f_file)
            .Add v
        Next v
    End With
    Set FileAsCollection = cll
    Set cll = Nothing

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

' Origin: Common Component "mVarTrans" =======================================
Private Function FileAsDictionary(ByVal f_file As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Returns a file's records/lines as Dictionary keys.
' Attention: Because the lines become Directory keys, they will become
'            distinct. I. e. each line will exist only once in the Dictionary.
'            To make this restriction productive, the number of occurrences
'            of each line is returned as item.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsDictionary"
    
    On Error GoTo eh
    Dim dct As New Dictionary
    Dim l   As Long
    Dim s   As String
    Dim v   As Variant
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    If Not FSo.FileExists(f_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named """ & f_file & """ does not exist!"
    
    With dct
        For Each v In FileAsArray(f_file)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .Item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set FileAsDictionary = KeySort(dct)
    Set dct = Nothing
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

' Origin: Common Component "mVarTrans" =======================================
Private Function FileAsFile(ByVal f_file_in As File, _
                            ByVal f_file_out As String, _
                   Optional ByVal f_append As Boolean = False, _
                   Optional ByVal f_rename As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns a file (f_file_in) as a file with another full name.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsFile"
    
    On Error GoTo eh
    With FSo
        Select Case True
            Case f_rename And Not f_append:     If f_file_in.Path = .GetParentFolderName(f_file_out) _
                                                Then f_file_in.Name = .GetFileName(f_file_out) _
                                                Else Err.Raise AppErr(1), ErrSrc(PROC), "File cannot be renamed when the provided file and the new file's name " & _
                                                                                        "do not point to the same location!"
            Case Not f_rename And Not f_append: .CopyFile f_file_in.Path, f_file_out
            Case Not f_rename And f_append:     StringAsFile FileAsString(f_file_in), f_file_out, True
            Case Else:                          Err.Raise AppErr(2), ErrSrc(PROC), "Rename  a n d  append is not supported!"
        End Select
    Set FileAsFile = .GetFile(f_file_out)
    End With
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

' Origin: Common Component "mVarTrans" =======================================
Private Function FileAsString(ByVal f_file As Variant, _
                     Optional ByRef f_split As String, _
                     Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns a file's (f_file) - provided as full name or object - records/lines
' as a single string with the records/lines delimited (f_split).
' ----------------------------------------------------------------------------
    Const PROC = "FileAsString"
    
    Dim s   As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    If Not FSo.FileExists(f_file) Then Err.Raise AppErr(1), ErrSrc(PROC), _
                                       "The file '" & f_file & "' provided does not exist!"
    
    Open f_file For Input As #1
    s = Input$(lOf(1), 1)
    Close #1
    
    Select Case True
        '~~ Return the line-break string to the caller enabling the transition of the string into an array
        Case InStr(s, vbCrLf) <> 0: f_split = vbCrLf
        Case InStr(s, vbCr) <> 0:   f_split = vbCr
        Case InStr(s, vbLf) <> 0:   f_split = vbLf ' e.g. the case when the file has been downloaded from GitHub
        Case Else
            Err.Raise AppErr(1), ErrSrc(PROC), "The line-break string in the file """ & f_file & """ is neither vbCrLf, nor vbCR, nor vbLf!"
    End Select
    
    '~~ Eliminate any trailing split string
    Do While Right(s, Len(f_split)) = f_split
        s = Left(s, Len(s) - Len(f_split))
        If Len(s) <= Len(f_split) Then Exit Do
    Loop
    
    If f_exclude_empty Then
        s = FileAsStringEmptyExcluded(s)
    End If
    FileAsString = s

xt: Exit Function

End Function

Private Function FileAsStringEmptyExcluded(ByVal s_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (s_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' ----------------------------------------------------------------------------
    
    s_s = FileStringTrimmed(s_s) ' leading and trailing empty already excluded
    Do While InStr(s_s, vbCrLf & vbCrLf) <> 0
        s_s = Replace(s_s, vbCrLf & vbCrLf, vbCrLf)
    Loop
    FileAsStringEmptyExcluded = s_s
    
End Function

Private Function FileFooterDelimiter(Optional ByVal d_length As Long = 2) As String
    FileFooterDelimiter = "; " & String(d_length, sFileFooterDelimiterChar)
End Function

Private Function FileHeaderDelimiter(Optional ByVal d_length As Long = 2) As String
    FileHeaderDelimiter = "; " & String(d_length, sFileHeaderDelimiterChar)
End Function

Private Function FileFooterDelimiterAdd(ByVal f_cll As Collection) As Collection
' ----------------------------------------------------------------------------
' Returns a provided collection of file footer-lines/comments with an added
' delimiter line in the length of the longest comment line.
' ----------------------------------------------------------------------------
    Dim l   As Long
    Dim v   As Variant
    Dim cll As New Collection
    
    If f_cll Is Nothing Then GoTo xt
    If f_cll.Count <= 1 Then GoTo xt
    FileFooterDelimiterRemove f_cll
    
    For Each v In f_cll
        If Left(v, 2) <> "; " Then v = "; " & v
'        cll.Add v
        l = Max(l, Len(v))
    Next v
    
    cll.Add FileFooterDelimiter(l - 2)
    For Each v In f_cll
        cll.Add v
    Next v
    Set FileFooterDelimiterAdd = cll
    
xt: Exit Function
End Function

Private Function FileFooterDelimiterRemove(ByRef f_cll As Collection) As Collection
    Dim l As Long
    Dim v As Variant
    
    If f_cll Is Nothing Then GoTo xt
    If f_cll.Count <= 1 Then GoTo xt
    If f_cll(1) Like FileFooterDelimiter(2) & "*" Then
        f_cll.Remove 1
    End If
    Set FileFooterDelimiterRemove = f_cll
    
xt: Exit Function
End Function

Private Function FileHeaderDelimiterAdd(ByVal f_cll As Collection) As Collection
' ----------------------------------------------------------------------------
' Returns a provided collection of file header lines/comments with an added
' delimiter line in the length of the longest comment line.
' ----------------------------------------------------------------------------
    Dim l   As Long
    Dim v   As Variant
    Dim cll As New Collection
    
    If f_cll Is Nothing Then GoTo xt
    If f_cll.Count <= 1 Then GoTo xt
    FileHeaderDelimiterRemove f_cll
    
    For Each v In f_cll
        If Left(v, 2) <> "; " Then v = "; " & v
        cll.Add v
        l = Max(l, Len(v))
    Next v
    
    cll.Add FileHeaderDelimiter(l - 2)
    Set FileHeaderDelimiterAdd = cll
    
xt: Exit Function
End Function

Private Function FileHeaderDelimiterRemove(ByRef f_cll As Collection) As Collection
    Dim l As Long
    Dim v As Variant
    
    If f_cll Is Nothing Then GoTo xt
    If f_cll.Count <= 1 Then GoTo xt
    If f_cll(f_cll.Count) Like FileHeaderDelimiter(2) & "*" Then
        f_cll.Remove f_cll.Count
    End If
    Set FileHeaderDelimiterRemove = f_cll
    
xt: Exit Function
End Function

Private Function FileStringTrimmed(ByVal s_s As String, _
                          Optional ByRef s_as_dict As Dictionary = Nothing) As String
' ----------------------------------------------------------------------------
' Returns a file as string (s_s) with any leading and trailing empty items,
' i.e. record, lines, excluded. When a Dictionary is provided
' the string is additionally returned as items with the line number as key.
' ----------------------------------------------------------------------------
    Dim s As String
    Dim i As Long
    Dim v As Variant
    
    s = s_s
    '~~ Eliminate any leading empty items
    Do While Left(s, 2) = vbCrLf
        s = Right(s, Len(s) - 2)
    Loop
    '~~ Eliminate a trailing eof if any
    If Right(s, 1) = VBA.Chr(26) Then
        s = Left(s, Len(s) - 1)
    End If
    '~~ Eliminate any trailing empty items
    Do While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Loop
    
    FileStringTrimmed = s
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

Public Sub HskpngNames(ByVal h_file As String, _
                  ParamArray h_hist() As Variant)
' ----------------------------------------------------------------------------
' Support the fact that a value-name may change over time. The argument is a
' list of strings, each specifying a value-name history:
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
' ----------------------------------------------------------------------------
    Const PROC = "HskpngNames"
    
    On Error GoTo eh
    Dim a           As Variant
    Dim aNames      As Variant
    Dim i           As Long
    Dim lFrom       As Long
    Dim lStep       As Long
    Dim lTo         As Long
    Dim sNameNew    As String
    Dim sNames      As String
    Dim sSects      As String
    Dim v           As Variant
    Dim dctHist     As Dictionary
    Dim dctValsInv  As Dictionary
    Dim dctSects    As Dictionary
    
    If Not IsValidFile(h_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(h_file)
    If Not FileIsActive(h_file) Then GoTo xt
    If Not ArrayIsAllocated(h_hist) Then GoTo xt
    
    Set dctHist = ArrayAsDictionary(h_hist)
    Set dctValsInv = DictionaryInverted(ValueNames(FileName, NAMES_HSKPNG_SECT_NAME))
    For Each v In dctValsInv
        If Not dctHist.Exists(v) Then ValueRemove dctValsInv(v), NAMES_HSKPNG_SECT_NAME
    Next v
    
    For Each v In dctHist
        If Not PPtwinHskpngNamesIsDone(h_file, v) Then
            sSects = vbNullString
            If InStr(v, ":") <> 0 Then
                '~~ Do names forwarding for specified section/s
                a = Split(v, ":")
                '~~ Isolate and unstrip names history
                sNames = a(UBound(a))
                sSects = Replace(v, ":" & sNames, vbNullString)
                sSects = Replace(sSects, ":", ",")
            Else
                '~~ Do names forwarding for all sections
                sNames = v
            End If
            
            Select Case True
                Case InStr(sNames, ">") <> 0 And InStr(v, "<") = 0
                    aNames = Split(sNames, ">")
                    sNameNew = aNames(UBound(aNames)) ' when vbNullString any old names are removed
                    lFrom = LBound(aNames)
                    lTo = UBound(aNames) - 1
                    lStep = 1
                Case InStr(sNames, "<") <> 0 And InStr(v, ">") = 0
                    aNames = Split(sNames, "<")
                    sNameNew = aNames(LBound(aNames)) ' when vbNullString any old names are removed
                    lFrom = UBound(aNames)
                    lTo = LBound(aNames) + 1
                    lStep = -1
                Case InStr(sNames, ">") <> 0 And InStr(v, "<") <> 0
                    Err.Raise AppErr(1), ErrSrc(PROC), "Ambigous history direction! Supported is either " < " or " > " but not both."
                Case Else
                    Err.Raise AppErr(2), ErrSrc(PROC), "No history direction provided (either " < " or " > ")."
            End Select
            
            If sSects = vbNullString Then sSects = DictionaryAsString(SectionNames(h_file))
            For i = lFrom To lTo Step lStep
                '~~ Loop through all history names and either delete them or move them forward to the current (new) name
                If sNameNew = vbNullString Then
                    PPtwinValueRemove h_file, sSects, aNames(i)
                Else
                    PPtwinValueRename h_file, sSects, aNames(i), sNameNew
                End If
                PPtwinHskpngNamesDone(h_file) = v
            Next i
        End If
    Next v
    PPmanage h_file

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
                 
Private Function IsComment(ByVal i_item As String) As Boolean
    IsComment = Len(Trim(i_item)) = 0 ' considers an empty line as a valid (separator) comment
    If Not IsComment Then
        IsComment = i_item Like ";*" And Not IsDelimiterFileHeader(i_item) And Not IsDelimiterFileFooter(i_item)
    End If
End Function

Private Function IsDelimiterFileFooter(ByVal i_item As String) As Boolean
    IsDelimiterFileFooter = i_item Like "; " & String(3, sFileFooterDelimiterChar) & "*"
End Function

Private Function IsDelimiterFileHeader(ByVal i_item As String) As Boolean
    IsDelimiterFileHeader = i_item Like "; " & String(3, sFileHeaderDelimiterChar) & "*"
End Function

Private Function IsEmptyFile(ByVal f_file_name As String, _
                    Optional ByRef f_file_arry As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the file (f_file) is empty, FALSE and the content as array
' otherwise.
' ----------------------------------------------------------------------------
    
    If FSo.FileExists(f_file_name) Then
        IsEmptyFile = FSo.GetFile(f_file_name).Size <= 2
    Else
        IsEmptyFile = True
    End If
    If Not IsEmptyFile Then
        f_file_arry = FileAsArray(f_file_name)
    End If
    
End Function

Private Function IsObject(ByVal i_var As Variant, _
                          ByRef i_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the object's (i_var) name (i_name) when a variant (i_var)
' is an object. When the object does not have a Name property an error is
' raised.
' ----------------------------------------------------------------------------
    Const PROC = "IsObject"
    
    If Not VBA.IsObject(i_var) Then Exit Function
    IsObject = True
    On Error Resume Next
    i_name = i_var.Name
    If Err.Number <> 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
         "VarTrans tries to use the Name property of an object when it is to be " & _
         "transferred into a string which is the case when String, Array, or File " & _
         "is the target format. However, the current item is an object which does " & _
         "not have a Name property!"
    
End Function

Public Function IsSystemSection(ByVal i_sect As String) As Boolean
    IsSystemSection = i_sect = NAMES_HSKPNG_SECT_NAME
End Function

Private Function IsValidFile(ByRef i_file As String, _
                    Optional ByRef i_arr As Variant, _
                    Optional ByRef i_valid As Boolean, _
                    Optional ByRef i_error As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a provided string (i_file) is a vaild Private Profile
' file, which is: - The file does not exist (but may be created)
'                 - or the file is empty
'                 - or the file is not empty and has no other lines but
'                   Section, Value, or Comment
' Sets up a Private Profile file twin when yet not already existing and up-to-
' date.
' Attention: This file validity check is restricted for being used by Public
'            methods and properties only!
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFile"
    
    On Error GoTo eh
    Dim sInvalidLine    As String
    
    If i_file = vbNullString And sFileName <> vbNullString Then
        '~~ The file has already been checked as valid
        i_file = sFileName
        IsValidFile = True
        If FSo.FileExists(i_file) Then
            i_valid = Not IsEmptyFile(i_file, i_arr)
            PPmanage i_file
            GoTo xt
        End If
     End If
     
    '~~ An explicitely specified file name is checked for validity and when valid
    '~~ becomes the FileName for all subsequent properties and methods which did not
    '~~ have a file name provided as argument.
    IsValidFile = True
    Select Case True
        Case Not IsValidFileContent(i_file, i_arr, sInvalidLine)
            '~~ The file is valid when it has a valid content (may be empty however)
            IsValidFile = False
            i_error = ErrInvalidFileContent(i_file, sInvalidLine)
        Case Else
            i_valid = True
            IsValidFile = True
    End Select
    If Not FSo.FileExists(i_file) Then
        PPtwinRemove i_file
    ElseIf IsValidFile = True Then
        PPmanage i_file
    End If

    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
    
End Function

Private Function IsValidFileContent(ByVal i_file_name As String, _
                                    ByRef i_file_arry As Variant, _
                                    ByRef i_line_nvld As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the file (i_file_name) either does not exist or contains
' only valid Private Profile lines else FALSE is returned and the invalid line
' (i_line_nvld).
' Note: Empty lines are a valid content.
' ----------------------------------------------------------------------------
    Const PROC = "IsValidFileContent"
    
    On Error GoTo eh
    Dim arr As Variant
    Dim i   As Long
    Dim s   As String
    
    If Not FSo.FileExists(i_file_name) Then
        IsValidFileContent = True ' an empty file is valid because it does not contain any invalid line
        GoTo xt
    End If
    
    arr = FileAsArray(i_file_name)
    IsValidFileContent = True
    For i = LBound(arr) To UBound(arr)
        s = arr(i)
        Select Case True
            Case IsValue(s)            ' Value
            Case PPfileIsSection(s)    ' Section
            Case Len(Trim(s)) = 0      ' Section separator
            Case Left(s, 1) = ";"      ' Comment or header/footer delimiter
            Case Else
                IsValidFileContent = False
                i_line_nvld = s
            GoTo xt
        End Select
    Next i
    i_file_arry = arr
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function IsValidNameSection(ByRef i_section As String) As Boolean
' ------------------------------------------------------------------------------
' Validate a section name on naming conventions in a Private Profile file
' Rules for are:
' - Must start with an alphanumeric character a _ or a @
' - May contain alphanumeric characters, hyphens, and underscores
' - Avoid spaces and special characters
' When no name is provided, the Property Section is used instead which had
' already been checked.
' ------------------------------------------------------------------------------
    Const PROC = "IsValidNameSection"
    
    On Error GoTo eh
    If i_section = vbNullString Then
        i_section = sSection
        IsValidNameSection = True
    Else
        With Reg
            .Global = True
            .IgnoreCase = True
            .Pattern = "^[A-Za-z0-9_@][A-Za-z0-9_@-]*$"
            IsValidNameSection = .Test(i_section)
        End With
    End If
   
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function IsValidNameValue(ByVal i_name As String) As Boolean
' ------------------------------------------------------------------------------
' Validate a value name/key on naming conventions in a Private Profile file
' Rules for are:
' - Must start with an alphanumeric character or a _
' - Can contain alphanumeric characters, hyphens, and underscores
' - Avoid spaces and special characters
' ------------------------------------------------------------------------------
    Const PROC = "IsValidNameValue"
    
    On Error GoTo eh
    IsValidNameValue = True

    With Reg
        .Global = True
        .IgnoreCase = True
        .Pattern = "^[A-Za-z0-9_][A-Za-z0-9_/-]*$"
        IsValidNameValue = .Test(i_name)
    End With
    
xt: Exit Function
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function IsValue(ByVal i_item As String, _
                Optional ByRef i_valu_name As String = vbNullString, _
                Optional ByRef i_valu As String = vbNullString) As Boolean
' ------------------------------------------------------------------------------
' When the item (i_item) is a Private Profile value name
' - and no value-name is provided, the function returns TRUE and the value-name
'   (i_valu_name)
' - and a value-name has been provided, the function only returns TRUE when the
'   item (i_item) is that value-name (i_value_name).
' ------------------------------------------------------------------------------
    Dim sName   As String
    Dim sValue  As String
    
    If i_item = vbNullString Then Exit Function
    If Left(i_item, 1) = "[" Or Left(i_item, 1) = ";" Then Exit Function
    
    sName = Split(i_item, "=")(0)
    sValue = Replace(i_item, sName & "=", vbNullString)
    If i_valu_name <> vbNullString Then
        If i_valu_name = sName Then
            IsValue = True
            i_valu = sValue
        End If
    Else
        IsValue = True
        i_valu_name = sName
        i_valu = sValue
    End If
    
End Function

Private Function KeySort(ByRef k_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (k_dct) sorted by key as Dictionary plus
' replacing the provided Dictionary (k_dct) by the sorted one.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr     As Variant
    Dim Temp    As Variant
    Dim i       As Long
    Dim j       As Long
    
    If k_dct Is Nothing Then GoTo xt
    If k_dct.Count = 0 Then GoTo xt
    
    With k_dct
        ReDim arr(0 To .Count - 1)
        For i = 0 To .Count - 1
            arr(i) = .Keys(i)
        Next i
    End With
    
    '~~ Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add vKey, k_dct.Item(vKey)
    Next i
    
xt: Set k_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

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

Private Sub PPfileAsPPtwin(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Sets up a Private Profile twin as a structure of Dictionaries and
' Collections which serves as a means to keep the content including its
' comments in ascending order, for existence checks, and for remove section
' and or values.
' ----------------------------------------------------------------------------
    Const PROC = "PPfileAsPPtwin"
    
    On Error GoTo eh
    Dim arr              As Variant
    Dim cll              As New Collection
    Dim i                As Long
    Dim s                As String
    Dim sName            As String
    Dim sValue           As String
    
    PPtwinRemove p_file
    arr = FileAsArray(p_file)
    Set cllPPtwinHeader = PPfileAsPPtwinPPfileHeader(p_file)
    Set cllPPtwinFooter = PPfileAsPPtwinPPfileFooter(p_file)
    Set dctPPtwinSections = New Dictionary
    Set dctPPtwinValues = Nothing
    
    For i = cllPPtwinHeader.Count To UBound(arr) - cllPPtwinFooter.Count
        sName = vbNullString
        s = vbNullString
        Select Case True
            Case IsComment(arr(i))
                '~~ Collect any comment except a file comment delimiter
                cll.Add arr(i)
            
            Case IsDelimiterFileHeader(arr(i))
                '~~ A file header delimiter indicates the end of a file header comments
                cll.Add arr(i)
                Set cllPPtwinHeader = cll
                Set cll = New Collection
            
            Case PPfileIsSection(arr(i), s)
                '~~ Stow section plus any possible comment (may be empty)
                If sSection <> vbNullString Then
                    '~~ This is a subsequent section which indicates that the
                    '~~ previous with its collected values has to be saved
                    PPfileAsPPtwinSectAdd p_sect:=sSection _
                                        , p_sect_comment:=cllPPtwinSectionComments
                    sSection = vbNullString
                End If
                sSection = s
                Set cllPPtwinSectionComments = cll
                Set cll = New Collection
            
            Case IsValue(arr(i), sName, sValue)
                '~~ Collect value in dctPPtwinValues with the value name as key
                PPfileAsPPtwinValueAdd p_value_name:=sName _
                                     , p_value:=sValue _
                                     , p_value_comments:=cll
        End Select
    Next i
    
    If dctPPtwinValues Is Nothing Then Set dctPPtwinValues = New Dictionary
    If dctPPtwinValues.Count > 0 Then
        If cll.Count <> 0 Then
            PPfileAsPPtwinSectAdd p_sect:=sSection _
                                , p_sect_comment:=cll
        Else
            PPfileAsPPtwinSectAdd p_sect:=sSection _
                                , p_sect_comment:=cllPPtwinSectionComments
        End If
        
        sSection = vbNullString
    End If
        
    '~~ Stow the file with possible comment and footer in the dctPPtwins Directory
    PPfileAsPPtwinAdd p_file
    
xt: Set cll = Nothing
    Set cllPPtwin = Nothing
    Set cllPPtwinHeader = Nothing
    Set dctPPtwinValues = Nothing
    Set cllPPtwinSection = Nothing
    Set cllPPtwinSectionComments = Nothing
    Set dctPPtwinValues = Nothing
    Set cllPPtwinValue = Nothing
    Set cllPPtwinValueComments = Nothing
    Set cllPPtwinFooter = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPfileAsPPtwinAdd(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Adds a new twin which had been setup from a Private Profile file (p_file) as
' Dictionary/collections structure, to the Directory of twins (dctPPtwins)
' with the file header (cllPPtwinHeader), file sections (dctPPtwinValues),
' and file footer (cllPPtwinFooter).
' ----------------------------------------------------------------------------
    Const PROC = "PPfileAsPPtwinAdd"
    
    On Error GoTo eh
    Dim cll     As New Collection
    
    If cllPPtwinHeader Is Nothing Then Set cllPPtwinHeader = New Collection
    If dctPPtwinSections Is Nothing Then Set dctPPtwinSections = New Dictionary
    If cllPPtwinFooter Is Nothing Then Set cllPPtwinFooter = New Collection
    
    cll.Add cllPPtwinHeader
    KeySort dctPPtwinSections
    PPfileAsPPtwinAddRemoveTopSectionSeparator
    cll.Add dctPPtwinSections
    cll.Add cllPPtwinFooter
    
    If dctPPtwins.Exists(p_file) Then dctPPtwins.Remove p_file
    dctPPtwins.Add p_file, cll
    Set cll = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPfileAsPPtwinAddRemoveTopSectionSeparator()
' ----------------------------------------------------------------------------
' Remove the top section's speparator line.
' Precondition: PPtwinSet had been processed
' ----------------------------------------------------------------------------
    Const PROC = "PPfileAsPPtwinAddRemoveTopSectionSeparator"
    
    On Error GoTo eh
    Dim sSect   As String
    
    If dctPPtwinSections.Count = 0 Then GoTo xt
    sSect = dctPPtwinSections.Keys()(0)                 ' get the top section's key
    PPtwinSectionSet sSect                              ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
    If cllPPtwinSectionComments(1) = vbNullString Then  ' remove a separator item if any
        cllPPtwinSectionComments.Remove 1
        PPtwinReplaceSection sSect                      ' update/re-write the twin section
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function PPfileAsPPtwinPPfileFooter(ByVal p_source As Variant) As Collection
' ----------------------------------------------------------------------------
' Returns a Private Profile file's footer comments as Collection which may be
' empty.
' ----------------------------------------------------------------------------
    Const PROC = "PPfileFooterGet"
    
    On Error GoTo eh
    Dim a   As Variant
    Dim cll As New Collection
    Dim i   As Long
    
    Set PPfileAsPPtwinPPfileFooter = cll
    Select Case VarType(p_source)
        Case vbArray:   a = p_source
        Case vbString:  a = FileAsArray(FSo.GetFile(p_source))
        Case Else:      Err.Raise AppErr(1), ErrSrc(PROC), "The provided source is neither an array nor a string which specifies an existing file!"
    End Select
    i = UBound(a)
    If i >= 0 Then
        Do While IsComment(a(i)) And i >= LBound(a)
            '~~ Any bottom comment up to the first non-comment is regarded a file footer
            cll.Add a(i)
            If i = LBound(a) Then Exit Do
            i = i - 1
        Loop
        If IsDelimiterFileFooter(a(i)) Then cll.Add a(i)
    End If
    Set PPfileAsPPtwinPPfileFooter = RevertOrder(cll)

xt: Set cll = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function PPfileAsPPtwinPPfileHeader(ByVal p_source As Variant) As Collection
    Const PROC = "PPfileAsPPtwinPPfileHeader"
    
    On Error GoTo eh
    Dim a   As Variant
    Dim cll As New Collection
    Dim i   As Long
    
    Set PPfileAsPPtwinPPfileHeader = cll
    Select Case VarType(p_source)
        Case vbArray:   a = p_source
        Case vbString:  a = FileAsArray(FSo.GetFile(p_source))
        Case Else:      Err.Raise AppErr(1), ErrSrc(PROC), "The provided source is neither an array nor a string which specifies an existing file!"
    End Select
    If UBound(a) >= 0 Then
        i = LBound(a)
        Do While IsComment(a(i)) And i <= UBound(a)
            '~~ Any top comment down to the first non-comment is regarded a file header
            cll.Add a(i)
            If i = UBound(a) Then Exit Do
            i = i + 1
        Loop
        If IsDelimiterFileHeader(a(i)) Then
            cll.Add a(i)
        End If
    End If
    Set PPfileAsPPtwinPPfileHeader = cll
    
xt: Set cll = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub PPfileAsPPtwinSectAdd(ByVal p_sect As String, _
                                  ByRef p_sect_comment As Collection)
' ----------------------------------------------------------------------------
' Adds a section (p_dct_sect) to the file Dictionary (p_file_dct).
' ----------------------------------------------------------------------------
    Const PROC = "PPfileAsPPtwinSectAdd"
    
    On Error GoTo eh
    Dim cll As New Collection
    
    If dctPPtwinSections Is Nothing Then Set dctPPtwinSections = New Dictionary
    If dctPPtwinValues Is Nothing Then Set dctPPtwinValues = New Dictionary
    If dctPPtwinValues.Count = 0 Then           ' empty section
        GoTo xt
    End If
    
    cll.Add PPfileAsPPtwinSectComments(p_sect_comment) ' transfer possible comments to the section Collection considering the separator option
    cll.Add KeySort(dctPPtwinValues)                   ' add the section as the last item to the Collection
    dctPPtwinSections.Add p_sect, cll                  ' add the Collection with the section name as key to the file Dictionary
    Set p_sect_comment = New Collection
    Set dctPPtwinValues = New Dictionary
    
xt: Set cll = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function PPfileAsPPtwinSectComments(ByVal p_cll As Collection) As Collection
' ----------------------------------------------------------------------------
' Sets up a sections comments by adding a top separator line when the
' respective option is TRUE.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim v   As Variant
    
    If bSectionsSeparated Then
        If p_cll.Count = 0 Then
            '~~ No comments at all yet
            cll.Add vbNullString
        ElseIf Len(Trim(p_cll(1))) <> 0 Then
            '~~ First item is not a sepeartor line
            cll.Add vbNullString
        End If
    End If
    For Each v In p_cll
        cll.Add v
    Next v
    Set PPfileAsPPtwinSectComments = cll
    Set cll = Nothing
    
End Function

Private Sub PPfileAsPPtwinValueAdd(ByVal p_value_name As String, _
                                   ByVal p_value As String, _
                                   ByRef p_value_comments As Collection)
' ----------------------------------------------------------------------------
' Add the first and only, the last one, or a subsequent value name
' (p_value_name), value (p_value), and possible comments (p_value_comments)
' to a section.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    If dctPPtwinValues Is Nothing Then Set dctPPtwinValues = New Dictionary
    
    cll.Add p_value_comments             ' add the value comment as item to the Collection
    cll.Add p_value                      ' add the value as item to the Collection
    If Not dctPPtwinValues.Exists(p_value_name) Then
        dctPPtwinValues.Add p_value_name, cll ' add the Collection
    End If
    Set p_value_comments = New Collection
    Set cll = Nothing
    
End Sub

Private Function PPfileHasChanged(ByVal p_file As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Private Profile file (p_file) has changed.
' ----------------------------------------------------------------------------
   PPfileHasChanged = StrComp(FileAsString(p_file) _
                            , dctPPtwinsAsString(p_file) _
                            , vbBinaryCompare) <> 0
End Function

Private Function PPfileIsSection(ByVal i_item As String, _
                        Optional ByRef i_sect_name As String = vbNullString) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the item (i_item) is a Private Profile section and it is
' equal to a provided section name (i_sect_name). When no section name
' (i_sect_name) is provided the function returns TRUE when the item (i_item)
' is a Private Profile Section.
' ------------------------------------------------------------------------------
    Dim sSection As String
    
    If Left(i_item, 1) = "[" And Right(i_item, 1) = "]" Then
        sSection = Replace(Replace(i_item, "[", vbNullString), "]", vbNullString)
        If i_sect_name <> vbNullString Then
            If sSection = i_sect_name Then
                PPfileIsSection = True
                
            End If
        Else
            PPfileIsSection = True
            i_sect_name = sSection
        End If
    End If
    
End Function

Public Sub PPmanage(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Manages a Private Profile by:
' - setup a Private Profile twin as a Dictionary structure when a twin doesn't
'   exist
' - re-setup a Private Profile twin in the form of a Dictionary structure when
'   the related Private Profile file has changed
' - re-write when the twin has changed (being an exception though).
' ----------------------------------------------------------------------------
    Const PROC = "PPmanage"
    
    On Error GoTo eh
    Dim bPPtwinHasChanged As Boolean
    Dim bPPfileHasChanged As Boolean
    
    Select Case True
        Case Not FSo.FileExists(p_file) And Not PPtwinExists(p_file)
            GoTo xt
        Case Not FSo.FileExists(p_file) And PPtwinExists(p_file)
            PPtwinAsPPfile p_file
            PPtwinClear
        Case Not PPtwinExists(p_file) And FSo.FileExists(p_file)
            PPfileAsPPtwin p_file
            PPtwinAsPPfile p_file ' ensure ascending order
            PPtwinClear
        Case PPtwinExists(p_file) And FSo.FileExists(p_file)
            bPPtwinHasChanged = PPtwinHasChanged(p_file)
            bPPfileHasChanged = PPfileHasChanged(p_file)
            Select Case True
                Case Not bPPtwinHasChanged And Not bPPfileHasChanged
                    '~~ Nothing (no change) to manage
                Case bPPtwinHasChanged And Not bPPfileHasChanged
                   '~~ re-write Private Profile file from the twin
                    PPtwinAsPPfile p_file
                    PPtwinClear
                Case bPPfileHasChanged And Not bPPtwinHasChanged
                    '~~ re-setup the twin to possibly re-organize items in ascending order
                    '~~ and replace the Private Profile file with the reorganized twin as string
                    PPfileAsPPtwin p_file
                    PPtwinAsPPfile p_file
                    PPtwinClear
                Case Else
                    Err.Raise AppErr(1), _
                    ErrSrc(PROC), "The twin has changed and the file has changed too, which is a severe logic error."
            End Select
    End Select

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

              
Private Sub PPtwinAsFile(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Writes a Private Profile's content based on a Private Profile's twin as
' string and saves the string for future change detection.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinAsFile"
    
    On Error GoTo eh
    Dim sPPtwin   As String
    
    BoP ErrSrc(PROC)
    sPPtwin = PPtwinAsString(p_file)
    StringAsFile sPPtwin, p_file
    
    If dctPPtwinsAsString.Exists(p_file) _
    Then dctPPtwinsAsString.Remove p_file
    dctPPtwinsAsString.Add p_file, sPPtwin
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPtwinAsPPfile(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Re-writes a Private Profile file (p_file) from its twin as string and saves
' the string for future change detections.
' ----------------------------------------------------------------------------
    Dim s As String
    
    s = PPtwinAsString(p_file)
    StringAsFile s, p_file
    
    If dctPPtwinsAsString.Exists(p_file) _
    Then dctPPtwinsAsString.Remove p_file
    dctPPtwinsAsString.Add p_file, s
    
End Sub

Private Function PPtwinAsString(ByVal p_file As String) As String
' ----------------------------------------------------------------------------
' Retuns a Private Profile's twin (p_file) transformed into a string.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinAsString"
    
    On Error GoTo eh
    Dim arrOut      As Variant
    Dim sSect       As String
    Dim vSect       As Variant
    Dim vName       As Variant
    Dim sName       As String
    
    If PPtwinExists(p_file) Then                                        ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        PPtwinAsStringComments arrOut, cllPPtwinHeader, True            ' File header
        
        For Each vSect In dctPPtwinSections
            sSect = vSect
            PPtwinSectionSet sSect                                      ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
            PPtwinAsStringComments arrOut, cllPPtwinSectionComments     ' Section comment
            Arry(arrOut) = "[" & sSect & "]"                            ' Section
            For Each vName In dctPPtwinValues
                sName = vName
                PPtwinValueSet sName                                    ' cllPPtwinValue, cllPPtwinValueComments, vPPtwinValue
                PPtwinAsStringComments arrOut, cllPPtwinValueComments   ' Value comment
                Arry(arrOut) = sName & "=" & vPPtwinValue               ' Value
            Next vName
        Next vSect
        
        PPtwinAsStringComments arrOut, cllPPtwinFooter, , True          ' File footer
        If ArrayIsAllocated(arrOut) Then
            PPtwinAsString = Join(arrOut, vbCrLf)
        End If
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub PPtwinAsStringComments(ByRef p_arr As Variant, _
                                   ByVal p_cll As Collection, _
                          Optional ByVal p_header As Boolean = False, _
                          Optional ByVal p_footer As Boolean = False)
    Dim v As Variant
    Dim s As String
    Dim l As Long
    Dim a As Variant
    
    If p_cll.Count > 0 Then
        If p_footer And Not IsDelimiterFileFooter(p_cll(1)) Then
            Arry(p_arr) = "; " & String(l - 2, sFileFooterDelimiterChar)
        End If
        For Each v In p_cll
            
            '~~ Accepts an empty line as a valid comment
            If v = vbNullString Then
                Arry(p_arr) = v
            Else
                s = Replace(";" & v, ";;", ";")
                l = Max(l, Len(s))
                Arry(p_arr) = s
            End If
        Next v
        If p_header And Not IsDelimiterFileHeader(p_arr(UBound(p_arr))) Then
            Arry(p_arr) = "; " & String(l - 2, sFileHeaderDelimiterChar)
        End If
    End If
    
End Sub

Private Sub PPtwinClear()
    Set cllPPtwin = Nothing
    Set cllPPtwinHeader = Nothing
    Set dctPPtwinSections = Nothing
    Set cllPPtwinSection = Nothing
    Set cllPPtwinSectionComments = Nothing
    Set dctPPtwinValues = Nothing
    Set cllPPtwinValue = Nothing
    Set cllPPtwinValueComments = Nothing
    Set cllPPtwinFooter = Nothing
End Sub

Private Function PPtwinExists(ByVal p_file As String) As Boolean
    If dctPPtwins.Exists(p_file) Then
        PPtwinExists = True
        PPtwinSet p_file ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
    End If
End Function

Private Function PPtwinFileFooterPrepare(ByVal f_comments As Collection) As Collection
' ----------------------------------------------------------------------------
' Returns comments (f_comments) as Collection representing a Private Profile
' file's footer - with a footer delimiter line at the top.
' ----------------------------------------------------------------------------

    Dim cll       As New Collection
    Dim lMax      As Long
    Dim v         As Variant
    
    If f_comments.Count > 0 Then
        '~~ Get max length footer comment
        For Each v In f_comments
            If Not IsDelimiterFileFooter(v) Then
                lMax = Max(lMax, Len(Comment(v)) - 2)
            End If
        Next v
        
        '~~ Prepare with footer delimiter
        cll.Add FileFooterDelimiter(lMax)
        For Each v In f_comments
            If Not IsDelimiterFileFooter(v) Then
                cll.Add Comment(v)
            End If
        Next v
    End If
    
    Set PPtwinFileFooterPrepare = cll
    Set cll = Nothing
    
End Function

Private Function PPtwinHasChanged(ByVal p_file As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Private Profile's twin (f_file) has changed.
' ----------------------------------------------------------------------------
    PPtwinHasChanged = StrComp(PPtwinAsString(p_file) _
                             , dctPPtwinsAsString(p_file) _
                             , vbBinaryCompare) <> 0
    
End Function

Private Function PPtwinHskpngNamesIsDone(ByVal h_file As String, _
                                         ByVal h_hist As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a name's (h_hist) housekeeping is done.
' Note: Each h_hist when done is registered under a NAMES_HSKPNG_VALUE_NAME
'       with a sequence number suffix.
' ----------------------------------------------------------------------------
    Dim s   As String
    Dim v   As Variant
    Dim dct As Dictionary
    Dim cll As Collection
    Dim sHist As String
    
    Set dct = PPtwinHskpngNamesNames(h_file)
    For Each v In dct
        Set cll = dct(v)
        sHist = cll(2)
        If sHist = h_hist Then
            PPtwinHskpngNamesIsDone = True
            Exit For
        End If
    Next v
        
End Function

Private Function PPtwinHskpngNamesNames(ByVal h_file As String) As Dictionary
    Dim dct As New Dictionary
    If PPtwinSectionExists(NAMES_HSKPNG_SECT_NAME, h_file) Then
        Set dct = dctPPtwinValues
    End If
    Set PPtwinHskpngNamesNames = dct
    Set dct = Nothing
End Function

                          
Private Function PPtwinNew(ByVal p_name As String, _
                           ByVal p_sect As String, _
                           ByVal p_valu As Variant) As Collection
' ----------------------------------------------------------------------------
' Adds a new Private Profile file's twin with one section (p_sect) and one
' value (p_name, p_valu).
' Note: There's no Private Profile file without at least on section with at
' one value (no values = no sections = no file).
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinNew"
    
    On Error GoTo eh
    Dim cll As New Collection
    
    If Not IsValidNameValue(p_name) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The new value's name does not meet the rules for value names in Provate Profile files!"
    If Not IsValidNameSection(p_sect) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The new section's name does not meet the rules for section names in a Private Profile file!" & vbLf & _
                                            "The name must start with an alpanumeric character, a _ or a @."
        
    Set cllPPtwin = New Collection
    cllPPtwin.Add cll ' live starts without a header
    Set dctPPtwinSections = New Dictionary
    dctPPtwinSections.Add p_sect, PPtwinNewSection(p_name, p_sect, p_valu)
    cllPPtwin.Add dctPPtwinSections
    cllPPtwin.Add cll ' live starts without a footer
    Set PPtwinNew = cllPPtwin
    PPtwinRemoveTopSectionSeparator dctPPtwinSections
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select

End Function

Private Function PPtwinNewSection(ByVal p_name As String, _
                                  ByVal p_sect As String, _
                                  ByVal p_valu As Variant) As Collection
' ----------------------------------------------------------------------------
' Adds a new value (p_name, p_valu) in a new section (p_sect) to a Private
' Profile file's twin (p_file) - provided the file/twin exists and the names
' for the value and the section conforms with the rules.
' Note: A new section allways comes with a new value
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinNewSection"
    
    On Error GoTo eh
    Dim cll As New Collection
    
    If Not IsValidNameValue(p_name) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The new value's name does not meet the rules for value names in Provate Profile files!"
    If Not IsValidNameSection(p_sect) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The new section's name does not meet the rules for section names in a Private Profile file!" & vbLf & _
                                            "The name must start with an alpanumeric character, a _ or a @."
    Set cllPPtwinSection = New Collection
    Set dctPPtwinValues = New Dictionary
    If bSectionsSeparated Then cll.Add vbNullString
    cllPPtwinSection.Add cll ' live starts without comments
    dctPPtwinValues.Add p_name, PPtwinNewValue(p_valu)
    cllPPtwinSection.Add dctPPtwinValues
    Set PPtwinNewSection = cllPPtwinSection
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

          
Private Function PPtwinNewValue(ByVal p_valu As Variant) As Collection
' ----------------------------------------------------------------------------
' Returns a Collection with the first item as the comments (p_comments) and
' the second item a the value (p_value).
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Set cllPPtwinValue = New Collection
    cllPPtwinValue.Add cll  ' Live starts without comments
    cllPPtwinValue.Add p_valu
    Set PPtwinNewValue = cllPPtwinValue
    
End Function

Private Sub PPtwinRemove(ByVal p_file As String)

    If dctPPtwins.Exists(p_file) Then dctPPtwins.Remove p_file
    If dctPPtwinsAsString.Exists(p_file) Then dctPPtwinsAsString.Remove p_file
    PPtwinClear
    
End Sub

Private Sub PPtwinRemoveTopSectionSeparator(ByVal p_sections As Dictionary)
' ----------------------------------------------------------------------------
' Remove the top section's speparator line.
' Precondition: PPtwinSet had been processed
' ----------------------------------------------------------------------------
    Const PROC = "PPfileAsPPtwinAddRemoveTopSectionSeparator"
    
    On Error GoTo eh
    Dim sSect   As String
    
    If p_sections.Count = 0 Then GoTo xt
    sSect = p_sections.Keys()(0)                    ' get the top section's key
    PPtwinSectionSet sSect                          ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
    If cllPPtwinSectionComments.Count > 0 Then
        If cllPPtwinSectionComments(1) = vbNullString _
        And bSectionsSeparated Then  ' remove a separator item if any
            cllPPtwinSectionComments.Remove 1
            PPtwinReplaceSection sSect                          ' update/re-write the twin section
        End If
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPtwinReplace(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Rmoves an existing Private Profile twin (p_file) and re-adds it - whereby
' any of the file's/twin's components may have changed.
' ----------------------------------------------------------------------------
    Dim cll As New Collection

    cll.Add cllPPtwinHeader
    cll.Add KeySort(dctPPtwinSections) ' in case a new section had been added
    cll.Add cllPPtwinFooter
    If dctPPtwins.Exists(p_file) Then dctPPtwins.Remove p_file
    dctPPtwins.Add p_file, cll
    Set cll = Nothing
    
End Sub

Private Sub PPtwinReplaceSection(ByVal p_section As String)
' ----------------------------------------------------------------------------
' Returns/replaces a Private Profile twin's section elements in the
' dtcPPtwinSections and arranges the sections in ascending order.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    cll.Add cllPPtwinSectionComments
    cll.Add KeySort(dctPPtwinValues) ' in case an new value had been added
    
    If dctPPtwinSections.Exists(p_section) Then dctPPtwinSections.Remove p_section
    dctPPtwinSections.Add p_section, cll
    
    KeySort dctPPtwinSections
    Set cll = Nothing

End Sub

                  
Private Sub PPtwinReplaceValue(ByVal p_name As String)
' ----------------------------------------------------------------------------
' Removes an exsting name (p_name) and re-adds it - whereby the comments and
' or the value may have changed
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    cll.Add cllPPtwinValueComments
    cll.Add vPPtwinValue
    If dctPPtwinValues.Exists(p_name) Then dctPPtwinValues.Remove p_name
    dctPPtwinValues.Add p_name, cll
    KeySort dctPPtwinValues
    Set cll = Nothing
    
End Sub

                                                                                
Private Function PPtwinSectionCommentsLet(ByVal p_comments As Collection) As Collection
    Dim cll As New Collection
    Dim v   As Variant
    
    If bSectionsSeparated Then cll.Add vbNullString
    For Each v In p_comments
        If cll.Count = 0 And v = vbNullString Then
        Else
            cll.Add Comment(v)
        End If
    Next v
    Set PPtwinSectionCommentsLet = cll
    
End Function

Private Function PPtwinSectionExists(ByVal p_sect As String, _
                                     ByVal p_file As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a section (p_sect) exists in the Private Profile file
' twin.
' ----------------------------------------------------------------------------
    If PPtwinExists(p_file) Then
        If dctPPtwinSections.Exists(p_sect) Then
            PPtwinSectionExists = True
            PPtwinSectionSet p_sect                 ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
        End If
    End If
    
End Function

Private Sub PPtwinSectionRemove(ByVal p_file As String, _
                                ByVal p_section As String)
' ----------------------------------------------------------------------------
' Removes an existing section (p_section) from a Private Profile file (p_file)
' twin.
' See PPfileAsPPtwin for the implemented structure.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinSectionRemove"
    
    On Error GoTo eh
    
    If PPtwinSectionExists(p_section, p_file) Then          ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
        dctPPtwinSections.Remove p_section
    End If
    
    Select Case True
        Case dctPPtwinSections Is Nothing
            dctPPtwins.Remove p_file
            dctPPtwinsAsString.Remove p_file
            FSo.DeleteFile p_file
        Case dctPPtwinSections.Count = 0
            '~~ When all sections had been removed
            '~~ the Private Profile file has become obsolete to and will be removed
            dctPPtwins.Remove p_file
            dctPPtwinsAsString.Remove p_file
            FSo.DeleteFile p_file
        Case dctPPtwinSections.Count = 1 And dctPPtwinSections.Keys(0) = NAMES_HSKPNG_SECT_NAME
            '~~ When the only section left is the names housekeepinfg section
            '~~ the Private Profile file has become obsolete to and will be removed
            dctPPtwins.Remove p_file
            dctPPtwinsAsString.Remove p_file
            FSo.DeleteFile p_file
        Case Else
            PPtwinReplace p_file            ' Replace the twin in the twins Directory
    End Select
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PPtwinSectionSet(ByVal p_section As String)
' ----------------------------------------------------------------------------
' Provides a Private Profile twin's section elements.
' ----------------------------------------------------------------------------
    Set cllPPtwinSection = dctPPtwinSections(p_section):    If cllPPtwinSection Is Nothing Then Set cllPPtwinSection = New Collection
    Set cllPPtwinSectionComments = cllPPtwinSection(1):     If cllPPtwinSectionComments Is Nothing Then Set cllPPtwinSectionComments = New Collection
    Set dctPPtwinValues = cllPPtwinSection(2):              If dctPPtwinValues Is Nothing Then Set dctPPtwinValues = New Collection

End Sub

Private Sub PPtwinSet(ByVal p_file As String)
' ----------------------------------------------------------------------------
' Provides a Private Profile twin's Header, Sections, and Footer
' ----------------------------------------------------------------------------
    Set cllPPtwin = dctPPtwins(p_file): If cllPPtwin Is Nothing Then Set cllPPtwin = New Collection
    Set cllPPtwinHeader = cllPPtwin(1): If cllPPtwinHeader Is Nothing Then Set cllPPtwinHeader = New Collection
    Set dctPPtwinSections = cllPPtwin(2)
    Set cllPPtwinFooter = cllPPtwin(3): If cllPPtwinFooter Is Nothing Then Set cllPPtwinFooter = New Collection
End Sub

Private Function PPtwinValueCommentsLet(ByVal p_comments As Collection) As Collection
    Dim cll As New Collection
    Dim v   As Variant
    
    For Each v In p_comments
        cll.Add Comment(v)
    Next v
    Set PPtwinValueCommentsLet = cll
    Set cll = Nothing
    
End Function

Private Function PPtwinValueExists(ByVal p_name As String, _
                                   ByVal p_sect As String, _
                                   ByVal p_file As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when a value (p_name) exists in a Private Profile file
' twin's section.
' ----------------------------------------------------------------------------
    
    If PPtwinExists(p_file) Then                    ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        If dctPPtwinSections.Exists(p_sect) Then
            PPtwinSectionSet p_sect                 ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
            If dctPPtwinValues.Exists(p_name) Then
                PPtwinValueExists = True
                PPtwinValueSet p_name               ' cllPPtwinValue, cllPPtwinValueComments, vPPtwinValue
            End If
        End If
    End If
    
End Function

Private Function PPtwinValueRemove(ByVal p_file As String, _
                                   ByVal p_sect As String, _
                                   ByVal p_name As String) As Boolean
' ----------------------------------------------------------------------------
' Removes from a Private Profile file (p_file) one or more values (p_name) in
' one or more sections (p_sections), whereby both may be specified as a comma
' separated string.
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinSectionRemove"
    
    On Error GoTo eh
    Dim sSect   As String
    Dim vSect   As Variant
    Dim vName   As Variant
    
    If p_sect <> vbNullString And p_name <> vbNullString Then
        If PPtwinExists(p_file) Then
            For Each vSect In StringAsArray(p_sect)
                sSect = vSect
                If PPtwinSectionExists(sSect, p_file) Then
                    PPtwinSectionSet sSect                              ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
                    For Each vName In StringAsArray(p_name)
                        If dctPPtwinValues.Exists(vName) Then
                            dctPPtwinValues.Remove vName    ' removes the value together with possible value comments
                            PPtwinValueRemove = True
                        End If
                    Next vName
                    
                    '~~ When all values in a section had been removed, the section becomes obsolete and is removed too
                    If dctPPtwinValues.Count > 0 Then
                        PPtwinReplaceSection sSect
                    Else
                        dctPPtwinSections.Remove sSect
                    End If
                End If
            Next vSect
        
            '~~ When all sections had been removed the Private Profile file has become obsolete to and should be removed
            Select Case True
                Case dctPPtwinSections Is Nothing
                    dctPPtwins.Remove p_file
                    dctPPtwinsAsString.Remove p_file
                    FSo.DeleteFile p_file
                Case dctPPtwinSections.Count = 0
                    dctPPtwins.Remove p_file
                    dctPPtwinsAsString.Remove p_file
                    FSo.DeleteFile p_file
                Case dctPPtwinSections.Count = 1 And dctPPtwinSections.Keys(0) = NAMES_HSKPNG_SECT_NAME
                    dctPPtwins.Remove p_file
                    dctPPtwinsAsString.Remove p_file
                    FSo.DeleteFile p_file
                Case Else
                    PPtwinReplace p_file            ' Replace the twin in the twins Directory
            End Select
        End If
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function PPtwinValueRename(ByVal p_file As String, _
                                   ByVal p_sect As String, _
                                   ByVal p_name_old As String, _
                                   ByVal p_name_new As String) As Boolean
' ----------------------------------------------------------------------------
' Renames in the Private Profile file twin (p_file) in the specified section/s
' (p_sect) any old name (p_name_old) by a new name (p_name_new).
' Returns TRUE when at least one name has been renamed
' ----------------------------------------------------------------------------
    Const PROC = "PPtwinValueRename"
    
    On Error GoTo eh
    Dim aSects  As Variant
    Dim sSect As String
    Dim v     As Variant
    
    If p_sect = vbNullString _
    Then aSects = DictionaryAsArray(SectionNames(p_file)) _
    Else aSects = StringAsArray(p_sect)
    
    If Not PPtwinExists(p_file) Then GoTo xt
    For Each v In aSects
        sSect = v
        If PPtwinValueExists(p_name_old, sSect, p_file) Then    ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
                                                                ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
                                                                ' cllPPtwinValue, cllPPtwinValueComments, vPPtwinValue
            dctPPtwinValues.Remove p_name_old
            PPtwinReplaceValue p_name_new
            PPtwinValueRename = True
            PPtwinReplaceSection sSect
            PPtwinReplace p_file
        End If
    Next v
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function PPtwinValues(ByVal p_file As String, _
                     Optional ByVal p_sections As String = vbNullString) As Dictionary
                     
    Dim dct         As New Dictionary
    Dim dctSections As Dictionary
    Dim vSect       As Variant
    Dim vValue      As Variant
    
    Set PPtwinValues = dct
    If PPtwinExists(p_file) Then                        ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        If p_sections = vbNullString _
        Then Set dctSections = dctPPtwinSections _
        Else Set dctSections = StringAsDictionary(p_sections)
        
        For Each vSect In dctPPtwinSections
            If dctSections.Exists(vSect) Then
                PPtwinSectionSet vSect                  ' cllPPtwinSection, cllPPtwinSectionComments, dctPPtwinValues
                For Each vValue In dctPPtwinValues
                    PPtwinValueSet vValue               ' cllPPtwinValue, cllPPtwinValueComments, vPPtwinValue
                    If Not dct.Exists(vValue) Then
                        dct.Add vValue, vPPtwinValue
                    End If
                Next vValue
            End If
        Next vSect
        Set PPtwinValues = dct
        Set dct = Nothing
    End If
                              
End Function

Private Sub PPtwinValueSet(ByVal p_name As String)
' ----------------------------------------------------------------------------
' Provides from a Private Profile file's twin the value elements.
' Precondition: The value is existing.
' ----------------------------------------------------------------------------
    Set cllPPtwinValue = dctPPtwinValues(p_name)
    Set cllPPtwinValueComments = cllPPtwinValue(1)
    vPPtwinValue = cllPPtwinValue(2)

End Sub

Private Function PPvalue(ByVal p_value As Variant, _
                Optional ByVal p_default As Variant) As Variant
' ----------------------------------------------------------------------------
' - When the option BooleanAsTrueFalseString = TRUE the function returns the
'   value (p_value) transformed into a TRUE/FALSE string when it is Boolean or
'   when it is a TRUE/FALSE string returns the corresponding boolean.
' - When the option BooleanAsTrueFalseString = FALSE the value is returned
'   converted into a string.
' ----------------------------------------------------------------------------
    Const PROC = "PPvalue"
    
    On Error GoTo eh
    
    If Not IsMissing(p_default) Then
        '~~ When a default had been provided, the string is returned converted
        '~~ to the value type of the default (typed value return)
        Select Case VarType(p_default)
            Case VarType(p_default) = vbBoolean
                If Not bBooleanAsTrueFalseString Then
                    '~~ Return the value as Boolean
                    PPvalue = VBA.CBool(p_value)
                Else
                    '~~ Consider the value is stored/returned as a True/False String
                    Select Case p_value
                        Case VALUE_TRUE:    PPvalue = True
                        Case VALUE_FALSE:   PPvalue = False
                        Case Else:          PPvalue = p_value
                    End Select
                End If
            
            Case VarType(p_default) = vbByte:       PPvalue = VBA.CByte(p_value)
            Case VarType(p_default) = vbCurrency:   PPvalue = VBA.CCur(p_value)
            Case VarType(p_default) = vbDate:       PPvalue = VBA.CDate(p_value)
            Case VarType(p_default) = vbDecimal:    PPvalue = VBA.CDec(p_value)
            Case VarType(p_default) = vbDouble:     PPvalue = VBA.CDbl(p_value)
            Case VarType(p_default) = vbInteger:    PPvalue = VBA.CInt(p_value)
            Case VarType(p_default) = vbLong:       PPvalue = VBA.CLng(p_value)
            Case VarType(p_default) = vbLongLong:   PPvalue = VBA.CLngLng(p_value)
            Case VarType(p_default) = vbSingle:     PPvalue = VBA.CSng(p_value)
            Case Else
                If Not bBooleanAsTrueFalseString Then
                    PPvalue = VBA.CStr(p_value)
                Else
                    '~~ Consider the value is stored/returned as a True/False String
                    Select Case p_value
                        Case VALUE_TRUE:    PPvalue = True
                        Case VALUE_FALSE:   PPvalue = False
                        Case Else:          PPvalue = p_value
                    End Select
                End If
        End Select
    Else
        '~~ When no default had been provided the value is returned
        '~~ considering the BooleanAsTrueFalseString option
        If bBooleanAsTrueFalseString Then
            Select Case p_value
                Case VALUE_TRUE:    PPvalue = True
                Case VALUE_FALSE:   PPvalue = False
                Case Else:          PPvalue = CStr(p_value)
            End Select
        Else
            PPvalue = CStr(p_value)
        End If
    End If ' Not IsMissing(p_default)
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    Const GITHUB_REPO_URL   As String = "https://github.com/warbe-maker/VBA-Private-Profile"
    
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        ShellRun GITHUB_REPO_URL & r_bookmark
    End If
        
End Sub

Private Function RevertOrder(ByVal r_cll As Collection) As Collection
' ----------------------------------------------------------------------------
' Returns a Collection (r_cll) with its items in reverse order.
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim i   As Long
    
    If r_cll.Count > 0 Then
        For i = r_cll.Count To 1 Step -1
            cll.Add r_cll(i)
        Next i
    End If
    Set RevertOrder = cll
    
End Function

Public Function SectionNames(Optional ByVal name_file As String, _
                             Optional ByRef names_cll As Collection, _
                             Optional ByRef names_dct As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns all section names in a Private Profile file (name_file) as
' Dictionary and optionally as Collection (names_cll) with all Section names
' in ascending order.
' ----------------------------------------------------------------------------
    Const PROC = "SectionNames"
    
    On Error GoTo eh
    Dim dct As New Dictionary
    Dim cll As New Collection
    Dim v   As Variant
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    
    If PPtwinExists(name_file) Then     ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        For Each v In dctPPtwinSections
            dct.Add v, vbNullString
            cll.Add v
        Next v
        Set names_cll = cll
        Set names_dct = dct
    End If
    Set SectionNames = dct
    
xt: Set dct = Nothing
    Set cll = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SectionRemove(ByVal name_section As String, _
                Optional ByVal name_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Removes one or more sections (name_section) in a Private Profile file
' (name_file) whereby several sections may be provided as a comma sepearated
' string.
' ----------------------------------------------------------------------------
    Const PROC = "SectionRemove"
    
    On Error GoTo eh
    Dim v As Variant
    
    If name_section = vbNullString Then GoTo xt
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If Not IsValidNameSection(name_section) Then Err.Raise AppErr(2), ErrSrc(PROC), ErrInvalidSectName(name_section)
    
    For Each v In StringAsArray(name_section)
        PPtwinSectionRemove name_file, v
    Next v
    PPmanage name_file
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function Sections(Optional ByVal name_file As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all sections in a Private Profile file (name_file)
' with the section name as the key.
' Note: In order to cover section comments the item is a Collection with two
'       items: 1. Collection of section comments
'              2. Dictionaray of values
' ----------------------------------------------------------------------------
    Const PROC = "Sections"
    
    Dim dct As New Dictionary
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If PPtwinExists(name_file) Then         ' cllPPtwin, cllPPtwinHeader, dctPPtwinSections, cllPPtwinFooter
        Set dct = dctPPtwinSections
    End If
    Set Sections = dct
    Set dct = Nothing
    
End Function

Private Function ShellRun(ByVal s_string As String, _
                 Optional ByVal s_show_how As Long = WIN_NORMAL) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
'                 - Call Email app: ShellRun("mailto:user@tutanota.com")
'                 - Open URL:       ShellRun("http://.......")
'                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
'                                   "Open With" dialog)
'                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
' Copyright:      This code was originally written by Dev Ashish. It is not to
'                 be altered or distributed, except as part of an application.
'                 You are free to use it in any application, provided the
'                 copyright notice is left unchanged.
' Courtesy of:    Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, s_string, vbNullString, vbNullString, s_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & s_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Private Function SplitIndctr(ByVal s_strng As String, _
                    Optional ByRef s_indctr As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns the split indicator of a string (s_strng) as string and as argument
' (s_indctr) provided no split indicator (s_indctr) is already provided.
' The dedection of a split indicator is bypassed in case one has already been
' provided.
' ----------------------------------------------------------------------------
    If s_indctr = vbNullString Then
        Select Case True
            Case InStr(s_strng, vbCrLf) <> 0: s_indctr = vbCrLf
            Case InStr(s_strng, vbLf) <> 0:   s_indctr = vbLf      ' e.g. in case of a downloaded file's_strng complete string
            Case InStr(s_strng, "|&|") <> 0:  s_indctr = "|&|"
            Case InStr(s_strng, ", ") <> 0:   s_indctr = ", "
            Case InStr(s_strng, "; ") <> 0:   s_indctr = "; "
            Case InStr(s_strng, ",") <> 0:    s_indctr = ","
            Case InStr(s_strng, ";") <> 0:    s_indctr = ";"
        End Select
    End If
    SplitIndctr = s_indctr

End Function

Private Function StringAsArray(ByVal s_items As String, _
                      Optional ByVal s_split As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    If s_items <> vbNullString Then
        If s_split = vbNullString Then s_split = SplitIndctr(s_items)
        StringAsArray = Split(s_items, s_split)
    End If

End Function

Private Function StringAsCollection(ByVal s_items As String, _
                           Optional ByVal s_split As String = vbNullString) As Collection
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    Dim v   As Variant
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    With cll
        For Each v In Split(s_items, s_split)
            .Add v
        Next v
    End With
    Set StringAsCollection = cll
    Set cll = Nothing
    
End Function

Private Function StringAsDictionary(ByVal s_items As String, _
                           Optional ByVal s_split As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Transforming the strings within a string (s_items) into a
'            Dictionary by saving the strings as key unifies them. As a
'            compensation of this restriction the number of occurences of a
'            string is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l   As Long
    Dim v   As Variant
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    
    With dct
        For Each v In Split(s_items, s_split)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                l = dct(v) + 1
                .Remove v
                .Add v, l
            End If
        Next v
    End With
    Set StringAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Private Function StringAsFile(ByVal s_strng As String, _
                     Optional ByRef s_file As Variant = vbNullString, _
                     Optional ByVal s_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes a string (s_strng) to a file (s_file) which might be a file object or
' a file's full name. When no file (s_file) is provided, a temporary file is
' returned. When a file's full name is provided any missing folder in the path
' is created on the fly
' Note : Only when the string has sub-strings delimited by vbCrLf the string
'        is written as records/lines.
' ----------------------------------------------------------------------------
    Dim sPath As String
    
    Select Case True
        Case s_file = vbNullString
            s_file = TempFile
        Case TypeName(s_file) = "File"
            s_file = s_file.Path
        Case Else
            sPath = FSo.GetParentFolderName(s_file)
            '~~ Create any missing folder
            CreatePath sPath
    End Select
    
    If s_file_append _
    Then Open s_file For Append As #1 _
    Else Open s_file For Output As #1
    Print #1, s_strng
    Close #1
    Set StringAsFile = FSo.GetFile(s_file)
    
End Function

Private Sub CreatePath(ByVal s_path As String)
' ----------------------------------------------------------------------------
' Creates any folder missing in a provided path.
' Precondition: The given drive exists and the calle has the right to create
'               any missing folder.
' ----------------------------------------------------------------------------
    Dim sParent As String

    If Not FSo.FolderExists(s_path) Then
        sParent = FSo.GetParentFolderName(s_path)
        '~~ Create missing parent folder recursively
        If Not FSo.FolderExists(sParent) _
        Then CreatePath sParent
        '~~ Create the current folder
        FSo.CreateFolder s_path
    End If
End Sub
Private Function StringAsString(ByVal s_items As String, _
                       Optional ByVal s_split As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a string (v_item) with any delimiter string replaced with vbCrLf.
' ----------------------------------------------------------------------------
    Const PROC = "StringAsString"
    
    On Error GoTo eh
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)

    StringAsString = Replace(s_items, s_split, vbCrLf)
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function TempFile(Optional ByVal f_path As String = vbNullString, _
                          Optional ByVal f_extension As String = ".txt") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(FSo.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = CurDir
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp
    FSo.CreateTextFile sTemp

End Function

Public Function ValueNameRename(ByVal name_value_old As String, _
                                ByVal name_value_new As String, _
                       Optional ByVal name_sections As Variant = vbNullString, _
                       Optional ByVal name_file As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Renames in a Private Profile file (name_file) a given value name
' (name_value_old) by a new name (name_value_new) in all provided
' sections (name_sections), when none are providedd, in all sections.
' ----------------------------------------------------------------------------
    Const PROC = "ValueNameRename"
    
    On Error GoTo eh
    Dim arr         As Variant
    Dim sSections   As Variant
    
    If name_value_old = vbNullString Or name_value_new = vbNullString Then GoTo xt
    If Not IsValidFile(name_file, arr) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
        
    If name_sections <> vbNullString _
    Then sSections = VarAsString(name_sections, ",") _
    Else sSections = DictionaryAsString(SectionNames(name_file), ",")
    
    If PPtwinValueRename(name_file, sSections, name_value_old, name_value_new) _
    Then PPmanage name_file
        
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function ValueNames(Optional ByVal name_file As String = vbNullString, _
                           Optional ByVal name_sections As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all value names a Private Profile file (name_file)
' with the value name as the key and the value as the item.
' When no section (name_sections) is specified the result is all value names in
' any section, if a section name is provided, only the value names is that
' section. When the file name (name_file) is ommitted it defaults to the name
' specified by the FileName property.
' Note: The returned value-names are distinct names! I.e. when a value exists
'       in more than one section it is still one distinct value-name.
' ----------------------------------------------------------------------------
    Const PROC = "ValueNames"
    
    On Error GoTo eh
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    Set ValueNames = PPtwinValues(name_file, name_sections)
    
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub ValueRemove(ByVal name_value As String, _
              Optional ByVal name_section As String = vbNullString, _
              Optional ByVal name_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Removes one or more values including a possible value comment by value-names
' (name_value) from a section (name_section) in a Private Profile file
' (name_file), whereby value-names may be provided as a comma separated string
' and section-names (name_section) may be provided as a comma separated string.
' Attention! When no section is provided, the value/value-name is removed in
' all sections it is used.
' When no file (name_file) is provided it defaults to the file name specified
' by the property FileName, when no section (name_section) is provided, the
' section defaults to the section name provided by the Section property.
' ----------------------------------------------------------------------------
    Const PROC = "ValueRemove"
    
    Dim sSections As String
    
    If Not IsValidFile(name_file) Then Err.Raise AppErr(1), ErrSrc(PROC), ErrInvalidFileName(name_file)
    If name_section = vbNullString _
    Then sSections = DictionaryAsString(SectionNames(name_file), ",") _
    Else sSections = VarAsString(name_section, ",")
    
    If PPtwinExists(name_file) Then
        If PPtwinValueRemove(name_file, sSections, name_value) _
        Then PPmanage name_file
    End If
xt:
End Sub

Private Function VarAsArray(ByVal v_items As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  VarAsArray = CollectionAsArray(v_items)
        Case "File":        VarAsArray = FileAsArray(v_items)
        Case "String":      If IsArray(v_items) _
                            Then VarAsArray = v_items _
                            Else VarAsArray = StringAsArray(v_items)
    End Select
    
End Function

Private Function VarAsCollection(ByVal v_items As Variant) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  Set VarAsCollection = v_items
        Case "File":        Set VarAsCollection = FileAsCollection(v_items)
        Case "String":      If IsArray(v_items) _
                            Then Set VarAsCollection = ArrayAsCollection(v_items) _
                            Else Set VarAsCollection = StringAsCollection(v_items)
    End Select
    
End Function

Private Function VarAsDictionary(ByVal v_items As Variant) As Dictionary
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  Set VarAsDictionary = CollectionAsDictionary(v_items)
        Case "File":        If Not v_items Is v_items _
                            Then Set VarAsDictionary = FileAsDictionary(v_items)
        Case "String":      If IsArray(v_items) _
                            Then Set VarAsDictionary = ArrayAsDictionary(v_items) _
                            Else Set VarAsDictionary = StringAsDictionary(v_items)
    End Select
    
End Function

Private Function VarAsFile(ByVal v_items As Variant, _
                  Optional ByVal v_file As String) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  Set VarAsFile = CollectionAsFile(v_items, v_file)
        Case "File":        If Not v_items Is v_items _
                            Then Set VarAsFile = FileAsFile(v_items, v_file)
        Case "String":      If IsArray(v_items) _
                            Then Set VarAsFile = ArrayAsFile(v_items, v_file) _
                            Else Set VarAsFile = StringAsFile(v_items, v_file)
    End Select
    
End Function

Private Function VarAsString(ByVal v_items As Variant, _
                    Optional ByVal v_delimiter As String = vbCrLf) As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Select Case TypeName(v_items)
        Case "Collection":  VarAsString = CollectionAsString(v_items, v_delimiter)
        Case "File":        VarAsString = FileAsString(v_items, vbCrLf)
        Case "String":      If IsArray(v_items) _
                            Then VarAsString = ArrayAsString(v_items, v_delimiter) _
                            Else VarAsString = StringAsString(v_items, v_delimiter)
    End Select
    
End Function

