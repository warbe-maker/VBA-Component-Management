Attribute VB_Name = "mComp"
Option Explicit
' ------------------------------------------------------------------------
' Standard-Module mComp
'                   Elementary services for a VBComponent either provided
'                   as object or name in in provided Workbbok.
'
' Public services:
' - Exists          Returns TRUE when a components name exists in a
'                   Workbook's VBProject
' - IsSheetDocMod   Returns TRUE when a provided VBComponent object is of
'                   a type Document-Module and represents a Worksheet
' - IsWrkbDocMod    Returns TRUE when a provided VBComponent object is of
'                   a type Document-Module and represents the Workbook
' - TempName        Returns a temporary name for a provided
'                   VBComponent's name which is not already used in
'                   the provided Workbook/VBProject
' - TypeString      Returns a provided VBComponent object's type as
'                   string
'
' ------------------------------------------------------------------------
Public Const RENAMED_BY_COMPMAN = "_RnmdByCompMan"

Public Enum vbcmType                ' Type of VBComponent
    vbext_ct_StdModule = 1          ' .bas
    vbext_ct_ClassModule = 2        ' .cls
    vbext_ct_MSForm = 3             ' .frm
    vbext_ct_ActiveXDesigner = 11   ' ??
    vbext_ct_Document = 100         ' .cls
End Enum

Public Function Exists(ByVal xst_wbk As Workbook, _
                       ByVal xst_vbc_name As String) As Boolean
' ----------------------------------------------------
' Returns TRUE when the component (comp_name) exists
' in the Workbook (wbk).
' -----------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = xst_wbk.VBProject.VBComponents(xst_vbc_name).Name
    Exists = Err.Number = 0
End Function

Public Function IsSheetDocMod(ByRef is_vbc As VBComponent, _
                     Optional ByRef is_wbk As Workbook = Nothing, _
                     Optional ByRef is_wsh_name As String = vbNullString) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Component (vbc) is the Worksheet object. When the
' optional Workbook (wbk) is provided, the sheet's Name is returned
' (sh_name).
' ------------------------------------------------------------------------
    Dim wsh As Worksheet
    
    IsSheetDocMod = is_vbc.Type = vbext_ct_Document And Not mComp.IsWrkbkDocMod(is_vbc)
    If Not is_wbk Is Nothing Then
        For Each wsh In is_wbk.Worksheets
            If wsh.CodeName = is_vbc.Name Then
                is_wsh_name = wsh.Name
                Exit For
            End If
        Next wsh
    End If
End Function

Public Function IsWrkbkDocMod(ByRef vbc As VBComponent) As Boolean
    
    Dim bSigned As Boolean
    On Error Resume Next
    bSigned = vbc.Properties("VBASigned").Value
    IsWrkbkDocMod = Err.Number = 0
    
End Function

Public Function TempName(ByVal tn_wbk As Workbook, _
                         ByVal tn_vbc_name As String) As String
' ----------------------------------------------------------------------------
' Returns a yet not existing temporary name for a component (tn_vbc_name).
' ----------------------------------------------------------------------------
    Dim i As Long
    
    TempName = tn_vbc_name & RENAMED_BY_COMPMAN
    Do
        On Error Resume Next
        TempName = tn_wbk.VBProject.VBComponents(TempName).Name
        If Err.Number <> 0 Then Exit Do ' a component with sTempName does not exist
        i = i + 1: TempName = TempName & i
    Loop
End Function

Public Function TypeString(ByVal vbc As VBComponent) As String
    Select Case vbc.Type
        Case vbext_ct_ActiveXDesigner:  TypeString = "ActiveX-Designer"
        Case vbext_ct_ClassModule:      TypeString = "Class-Module"
        Case vbext_ct_Document
            If mComp.IsSheetDocMod(vbc) _
            Then TypeString = "Document-Module (Worksheet)" _
            Else TypeString = "Document-Module (Workbook)"
        Case vbext_ct_MSForm:           TypeString = "UserForm"
        Case vbext_ct_StdModule:        TypeString = "Standard-Module"
    End Select
End Function

