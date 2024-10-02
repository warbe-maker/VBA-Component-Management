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
Private Const VALUE_FONT_SIZE           As String = "Font/PointSize"
Private Const VALUE_BAR0                As String = "Settings-Bar0/Visible"
Private Const VALUE_BAR1                As String = "Settings-Bar1/Visible"
Private Const SECTION_NAME              As String = "WinMerge"

Private PrivProf                        As clsPrivProf

Private Property Let Value(Optional ByVal pp_value_name As String, _
                                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes a value (pp_value) named (pp_value_name) into file (pp_file).
' ----------------------------------------------------------------------------
    Const PROC = "Value"
    
    On Error GoTo eh
    If PrivProf Is Nothing Then
        Set PrivProf = New clsPrivProf
        PrivProf.FileName = WinMergeIniFullName
        PrivProf.Section = SECTION_NAME
    End If
    PrivProf.Value(pp_value_name) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get WinMergeIniFullName() As String:        WinMergeIniFullName = ThisWorkbook.Path & "\WinMerge.ini":      End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mWinMergeIni." & sProc
End Function

Public Sub Setup(ByVal s_ini_file As String)
' ----------------------------------------------------------------------------
' CompMan only writes the required options. When WinMerge is executed for the
' first time it will write all the remaining properties.
' ----------------------------------------------------------------------------
    Value(VALUE_BAR0) = 0
    Value(VALUE_BAR1) = 0
    Value(VALUE_FONT_SIZE) = 10
    Value(VALUE_NAME_IGNORE_BLANKS) = 1
    Value(VALUE_NAME_IGNORE_CASE) = 1
End Sub

