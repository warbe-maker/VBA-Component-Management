Attribute VB_Name = "mPending"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mRaw
'                 Maintains for raw components identified by their
'                 component name the values WbkFullName and ExpFileFullName.
' ---------------------------------------------------------------------------

Private Const VALUE_HOST_FULL_NAME = "WbkFullName"
Private Const VALUE_EXP_FILE_FULL_NAME = "ExpFileFullName"
Private Const VALUE_COMP_TYPE = "ComponentType"
Private sServiced As String

Public Property Let Serviced(ByVal sv_wb As Workbook): sServiced = sv_wb.PATH & "\Pending.imp":     End Property

Public Property Let Import( _
            Optional ByVal im_wb As Workbook, _
            Optional ByVal im_exp_file_full_name As String, _
            Optional ByVal im_comp_type As Long, _
                     ByVal im_comps As Variant)
' -----------------------------------------------------------
' Write a pending import component to the Pending.imp file
' -----------------------------------------------------------
    ExpFileFullName(comp_name:=im_comps) = im_exp_file_full_name
    WbkFullName(comp_name:=im_comps) = im_wb.FullName
    CompType(ct_comp_name:=im_comps) = im_comp_type
End Property

Private Property Get Import( _
            Optional ByVal im_wb As Workbook, _
            Optional ByVal im_exp_file_full_name As String, _
            Optional ByVal im_comp_type As Long) As Variant
' -----------------------------------------------------------------
' Read all pending import components and return them as Dictionary.
' -----------------------------------------------------------------
    Set Import = CompNames
End Property

Private Property Get ExpFileFullName( _
                     Optional ByVal comp_name As String) As String
    ExpFileFullName = value(vl_section:=comp_name, vl_value_name:=VALUE_EXP_FILE_FULL_NAME)
End Property

Private Property Let ExpFileFullName( _
                     Optional ByVal comp_name As String, _
                              ByVal ef_full_name As String)
    value(vl_section:=comp_name, vl_value_name:=VALUE_EXP_FILE_FULL_NAME) = ef_full_name
End Property

Public Property Get CompType(Optional ByVal ct_comp_name As String) As vbcmType
    CompType = value(vl_section:=ct_comp_name, vl_value_name:=VALUE_COMP_TYPE)
End Property

Public Property Let CompType(Optional ByVal ct_comp_name As String, _
                                      ByVal ct_comp_type As vbcmType)
    value(vl_section:=ct_comp_name, vl_value_name:=VALUE_COMP_TYPE) = ct_comp_type
End Property

Private Property Get WbkFullName( _
                     Optional ByVal comp_name As String) As String
    WbkFullName = value(vl_section:=comp_name, vl_value_name:=VALUE_HOST_FULL_NAME)
End Property

Public Property Let WbkFullName( _
                     Optional ByVal comp_name As String, _
                              ByVal hst_full_name As String)
    value(vl_section:=comp_name, vl_value_name:=VALUE_HOST_FULL_NAME) = hst_full_name
End Property

Private Property Get value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String) As Variant
    
    value = mFile.value(vl_file:=sServiced _
                      , vl_section:=vl_section _
                      , vl_value_name:=vl_value_name _
                       )
End Property

Private Property Let value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String, _
                    ByVal vl_value As Variant)
' --------------------------------------------------
' Write the value (vl_value) named (vl_value_name)
' into the file RAWS_DAT_FILE.
' --------------------------------------------------
    
    mFile.value(vl_file:=sServiced _
              , vl_section:=vl_section _
              , vl_value_name:=vl_value_name _
               ) = vl_value

End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mPending" & "." & sProc
End Function


Public Sub Resolve()
' ------------------------------------------------------------------
' Resolve all pending imports provided there are one and clear the
' Pending.imp file.
' Preconditions when not met the service terminates without notice:
' - The component's Workbook is open
' - The component doe not exist
' - The Export File exists
'
' ------------------------------------------------------------------
    Const PROC = "Resolve"
    
    On Error GoTo eh
    Dim fso                 As New FileSystemObject
    Dim wb                  As Workbook
    Dim sWbkFullName        As String
    Dim sExpFileFullName    As String
    Dim v                   As Variant
    
    If Not fso.FileExists(sServiced) Then GoTo xt
    For Each v In CompNames
        sWbkFullName = WbkFullName(comp_name:=v)
        sExpFileFullName = ExpFileFullName(comp_name:=v)
        
        If Not mCompMan.WbkIsOpen(io_full_name:=sWbkFullName) Then GoTo xt
        If Not fso.FileExists(sExpFileFullName) Then GoTo xt
        
        Set wb = mCompMan.WbkGetOpen(go_wb_full_name:=sWbkFullName)
        If mCompMan.CompExists(ce_wb:=wb, ce_comp_name:=v) Then GoTo xt
        
        wb.VBProject.VBComponents.Import sExpFileFullName
        wb.VBProject.VBComponents(v).Export sExpFileFullName
    Next v

    If fso.FileExists(sServiced) Then fso.DeleteFile (sServiced)
    
xt: Set cComp = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Function CompNames() As Dictionary
    Set CompNames = mFile.SectionNames(sn_file:=sServiced)
End Function

Public Sub Remove(ByVal comp_name As String)
    mFile.SectionsRemove sr_file:=sServiced _
                       , sr_section_names:=comp_name
End Sub


