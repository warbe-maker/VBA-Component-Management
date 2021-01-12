Attribute VB_Name = "mRaw"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mRaw
'                 Maintains for raw components identified by their
'                 component name the values HostFullName and ExpFileFullName.
' ---------------------------------------------------------------------------

Private Const VALUE_HOST_FULL_NAME      As String = "HostFullName"
Private Const VALUE_EXP_FILE_FULL_NAME  As String = "ExpFileFullName"

Public Property Get DAT_FILE() As String: DAT_FILE = mMe.CompManAddinPath & "\Raws.dat":   End Property

Public Property Get ExpFileFullName( _
                     Optional ByVal comp_name As String) As String
    ExpFileFullName = value(vl_section:=comp_name, vl_value_name:=VALUE_EXP_FILE_FULL_NAME)
End Property

Public Property Let ExpFileFullName( _
                     Optional ByVal comp_name As String, _
                              ByVal ef_full_name As String)
    value(vl_section:=comp_name, vl_value_name:=VALUE_EXP_FILE_FULL_NAME) = ef_full_name
End Property

Public Property Get HostFullName( _
                     Optional ByVal comp_name As String) As String
    HostFullName = value(vl_section:=comp_name, vl_value_name:=VALUE_HOST_FULL_NAME)
End Property

Public Property Let HostFullName( _
                     Optional ByVal comp_name As String, _
                              ByVal hst_full_name As String)
    value(vl_section:=comp_name, vl_value_name:=VALUE_HOST_FULL_NAME) = hst_full_name
End Property

Private Property Get value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String) As Variant
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    value = mFile.value(vl_file:=DAT_FILE _
                      , vl_section:=vl_section _
                      , vl_value_name:=vl_value_name _
                       )
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: Exit Property
    End Select
End Property

Private Property Let value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String, _
                    ByVal vl_value As Variant)
' --------------------------------------------------
' Write the value (vl_value) named (vl_value_name)
' into the file RAWS_DAT_FILE.
' --------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFile.value(vl_file:=DAT_FILE _
              , vl_section:=vl_section _
              , vl_value_name:=vl_value_name _
               ) = vl_value

xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: Exit Property
    End Select
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRaws" & "." & sProc
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
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Function Components() As Dictionary
    Set Components = mFile.SectionNames(sn_file:=DAT_FILE)
End Function

Public Sub Remove(ByVal comp_name As String)
    mFile.SectionsRemove sr_file:=DAT_FILE _
                       , sr_section_names:=comp_name
End Sub


