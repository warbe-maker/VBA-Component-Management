Attribute VB_Name = "mWinMergeIni"
Option Explicit
' ----------------------------------------------------------------------------
' Provides the WinMerge.ini in ThisWorkbook's parent folder for being used
' when displaying code differences.
' [WinMerge]
' Settings/IgnoreBlankLines=1
' Settings/IgnoreCase=1
' ----------------------------------------------------------------------------
Private Const VALUE_NAME_IGNORE_BLANKS  As String = "Settings/IgnoreBlankLines"
Private Const VALUE_NAME_IGNORE_CASE    As String = "Settings/IgnoreCase"
Private Const SECTION_NAME              As String = "WinMerge"

Private Property Let Value(Optional ByVal pp_value_name As String, _
                           Optional ByVal pp_file As String, _
                                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes a value (pp_value) named (pp_value_name) into file (pp_file).
' ----------------------------------------------------------------------------
    Const PROC = "Value"
    
    On Error GoTo eh
    mFso.PPvalue(pp_file:=pp_file _
              , pp_section:=SECTION_NAME _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get WinMergeIniAddinFullName() As String:   WinMergeIniAddinFullName = ThisWorkbook.Path & "\Addin\WinMerge.ini":   End Property

Public Property Get WinMergeIniFullName() As String:        WinMergeIniFullName = ThisWorkbook.Path & "\WinMerge.ini":              End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mWinMergeIni." & sProc
End Function

Public Sub Setup(ByVal s_ini_file As String)
' ----------------------------------------------------------------------------
' CompMan only writes the required options. When WinMerge is executed for the
' first time it will write all the remaining properties.
'
' ----------------------------------------------------------------------------
    Value(VALUE_NAME_IGNORE_BLANKS, s_ini_file) = 1
    Value(VALUE_NAME_IGNORE_CASE, s_ini_file) = 1
End Sub

