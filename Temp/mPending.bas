Attribute VB_Name = "mPending"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mRaw
'                 Maintains for raw components identified by their
'                 component name the values WbkFullName and ExpFilePath.
' ---------------------------------------------------------------------------
Private Const VALUE_HOST_FULL_NAME = "WbkFullName"
Private Const VALUE_EXP_FILE_FULL_NAME = "ExpFilePath"
Private Const VALUE_COMP_TYPE = "ComponentType"

Private sPendingFileFullName As String

Public Property Get PendingFileFullName(ByVal pf_wb As Workbook)
    PendingFileFullName = mMe.PendingServicesFile(pf_wb)
End Property

Public Property Let Import( _
            Optional ByVal im_wb As Workbook, _
            Optional ByVal im_exp_file_full_name As String, _
            Optional ByVal im_comp_type As Long, _
                     ByVal im_comps As Variant)
' -----------------------------------------------------------
' Write a pending import component to the Pending.imp file
' -----------------------------------------------------------
    mPending.ServicedWbk = im_wb
    ExpFilePath(comp_name:=im_comps) = im_exp_file_full_name
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
    mPending.ServicedWbk = im_wb
    Set Import = CompsName
End Property

Private Property Get ExpFilePath( _
                     Optional ByVal comp_name As String) As String
    ExpFilePath = Value(vl_section:=comp_name, vl_value_name:=VALUE_EXP_FILE_FULL_NAME)
End Property

Private Property Let ExpFilePath( _
                     Optional ByVal comp_name As String, _
                              ByVal ef_full_name As String)
    Value(vl_section:=comp_name, vl_value_name:=VALUE_EXP_FILE_FULL_NAME) = ef_full_name
End Property

Public Property Get CompType(Optional ByVal ct_comp_name As String) As vbcmType
    CompType = Value(vl_section:=ct_comp_name, vl_value_name:=VALUE_COMP_TYPE)
End Property

Public Property Let CompType(Optional ByVal ct_comp_name As String, _
                                      ByVal ct_comp_type As vbcmType)
    Value(vl_section:=ct_comp_name, vl_value_name:=VALUE_COMP_TYPE) = ct_comp_type
End Property

Private Property Get WbkFullName( _
                     Optional ByVal comp_name As String) As String
    WbkFullName = Value(vl_section:=comp_name, vl_value_name:=VALUE_HOST_FULL_NAME)
End Property

Public Property Let WbkFullName( _
                     Optional ByVal comp_name As String, _
                              ByVal hst_full_name As String)
    Value(vl_section:=comp_name, vl_value_name:=VALUE_HOST_FULL_NAME) = hst_full_name
End Property

Private Property Get Value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String) As Variant
    
    Value = mFile.Value(vl_file:=sPendingFileFullName _
                      , vl_section:=vl_section _
                      , vl_value_name:=vl_value_name _
                       )
End Property

Private Property Let Value( _
           Optional ByVal vl_section As String, _
           Optional ByVal vl_value_name As String, _
                    ByVal vl_value As Variant)
' --------------------------------------------------
' Write the value (vl_value) named (vl_value_name)
' into the file RAWS_DAT_FILE.
' --------------------------------------------------
    
    mFile.Value(vl_file:=sPendingFileFullName _
              , vl_section:=vl_section _
              , vl_value_name:=vl_value_name _
               ) = vl_value

End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mPending" & "." & sProc
End Function

Public Function Still(ByVal st_comp_name As String) As Boolean
    Still = CompsName.Exists(st_comp_name)
End Function

Public Property Let ServicedWbk(ByVal wb As Workbook)
    sPendingFileFullName = mMe.PendingServicesFile(wb)
End Property

Public Sub Resolve(ByVal rs_wb As Workbook)
' -----------------------------------------------------
' Resolve all pending imports for the Workbook (rs_wb).
' Preconditions when not met the service terminates
' without notice:
' - The component's Workbook is open
' - The component doe not exist
' - The Export File exists
' -----------------------------------------------------
    Const PROC = "Resolve"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim wb              As Workbook
    Dim sWbkFullName    As String
    Dim sExpFilePath    As String
    Dim v               As Variant
    Dim cPendingComp    As New clsComp
    Dim sComp           As String
    
    mPending.ServicedWbk = rs_wb
    If Not fso.FileExists(sPendingFileFullName) Then GoTo xt
    
    With cPendingComp
        For Each v In mPending.CompsName
            sComp = v
            .WrkbkFullName = WbkFullName(comp_name:=sComp)
            .CompName = sComp
            sExpFilePath = ExpFilePath(comp_name:=sComp)
            
            If Not mCompMan.WbkIsOpen(io_full_name:=.WrkbkFullName) Then GoTo xt
            If Not fso.FileExists(.ExpFilePath) Then GoTo xt
            
            Set wb = mCompMan.WbkGetOpen(go_wb_full_name:=.WrkbkFullName)
            
            '~~ When the pending import component already exists the entry is removed
            '~~ from the pending file and the service is terminated without notice.
            If mCompMan.CompExists(ce_wb:=wb, ce_comp_name:=sComp) Then
                Remove sComp
                GoTo xt
            End If
            
            '~~ Only when the import has not failed the pending import entry
            '~~ is removed and the component is exported when the imported
            '~~ Export File is not the component's origin Export File
            On Error Resume Next
            wb.VBProject.VBComponents.Import sExpFilePath
            If Err.Number = 0 Then
                Remove sComp
                If sExpFilePath <> .ExpFilePath _
                Then wb.VBProject.VBComponents(v).Export sExpFilePath
            End If
        Next v

    End With
    
xt: If mPending.CompsName.Count = 0 And fso.FileExists(sPendingFileFullName) Then fso.DeleteFile (sPendingFileFullName)
    Set cPendingComp = Nothing
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Function CompsName() As Dictionary
    
    Dim fso As New FileSystemObject
    Dim dct As New Dictionary
    
    If fso.FileExists(sPendingFileFullName) _
    Then Set dct = mFile.SectionNames(sn_file:=sPendingFileFullName)

xt: Set CompsName = dct
    Set fso = Nothing

End Function

Public Sub Remove(ByVal comp_name As String)
    mFile.SectionsRemove sr_file:=sPendingFileFullName _
                       , sr_section_names:=comp_name
End Sub


