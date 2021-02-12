Attribute VB_Name = "mHostedRaws"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mRaw
'                 Maintains for raw components identified by their
'                 component name the values HostFullName and ExpFileFullName.
' ---------------------------------------------------------------------------

Private Const VNAME_HOST_FULL_NAME      As String = "HostFullName"
Private Const VNAME_EXP_FILE_FULL_NAME  As String = "ExpFileFullName"

Public Property Get DAT_FILE() As String: DAT_FILE = mMe.CompManAddinPath & "\HostedRaws.dat":   End Property

Public Property Get ExpFileFullName( _
                     Optional ByVal comp_name As String) As String
    ExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_EXP_FILE_FULL_NAME)
End Property

Public Property Let ExpFileFullName( _
                     Optional ByVal comp_name As String, _
                              ByVal ef_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_EXP_FILE_FULL_NAME) = ef_full_name
End Property

Public Property Get HostFullName( _
                     Optional ByVal comp_name As String) As String
    HostFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_HOST_FULL_NAME)
End Property

Public Property Let HostFullName( _
                     Optional ByVal comp_name As String, _
                              ByVal hst_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_HOST_FULL_NAME) = hst_full_name
End Property

Private Property Get Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String) As Variant
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFile.Value(pp_file:=DAT_FILE _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: Exit Property
    End Select
End Property

Private Property Let Value( _
           Optional ByVal pp_section As String, _
           Optional ByVal pp_value_name As String, _
                    ByVal pp_value As Variant)
' --------------------------------------------------
' Write the value (pp_value) named (pp_value_name)
' into the file RAWS_DAT_FILE.
' --------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFile.Value(pp_file:=DAT_FILE _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: Exit Property
    End Select
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRaw" & "." & sProc
End Function

Public Function Exists(ByVal raw_comp_name As String) As Boolean
    Exists = Components.Exists(raw_comp_name)
End Function

Public Function MaxRawLenght() As Long
' -----------------------------------------------
' Returns the max length of a raw componen's name
' -----------------------------------------------
    Const PROC = "MaxRawLenght"
    
    On Error GoTo eh
    Dim v As Variant
    
    For Each v In Components
        MaxRawLenght = Max(MaxRawLenght, Len(v))
    Next v
    
xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Function Components() As Dictionary
    Set Components = mFile.SectionNames(DAT_FILE)
End Function

Public Sub Remove(ByVal comp_name As String)
    mFile.SectionsRemove pp_file:=DAT_FILE _
                       , pp_sections:=comp_name
End Sub


