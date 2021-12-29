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
Public Const RENAMED_BY_COMPMAN = "_RenamedByCompMan"

Public Function Exists( _
                 ByRef wb As Workbook, _
                 ByVal comp_name As String) As Boolean
' ----------------------------------------------------
' Returns TRUE when the component (comp_name) exists
' in the Workbook (wb).
' -----------------------------------------------------
    Dim s As String
    On Error Resume Next
    s = wb.VBProject.VBComponents(comp_name).Name
    Exists = Err.Number = 0
End Function

Public Function IsSheetDocMod( _
                        ByRef vbc As VBComponent, _
               Optional ByRef wb As Workbook = Nothing, _
               Optional ByRef sh_name As String = vbNullString) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Component (vbc) is the Worksheet object. When the
' optional Workbook (wb) is provided, the sheet's Name is returned
' (sh_name).
' ------------------------------------------------------------------------
    Dim ws As Worksheet
    
    IsSheetDocMod = vbc.Type = vbext_ct_Document And Not mComp.IsWrkbkDocMod(vbc)
    If Not wb Is Nothing Then
        For Each ws In wb.Worksheets
            If ws.CodeName = vbc.Name Then
                sh_name = ws.Name
                Exit For
            End If
        Next ws
    End If
End Function

Public Function IsWrkbkDocMod(ByRef vbc As VBComponent) As Boolean
    
    Dim bSigned As Boolean
    On Error Resume Next
    bSigned = vbc.Properties("VBASigned").Value
    IsWrkbkDocMod = Err.Number = 0
    
End Function

Public Function TempName(ByVal tn_wb As Workbook, _
                         ByVal tn_comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns a yet not existing temporary name for a component (tn_comp_name).
' ----------------------------------------------------------------------------
    Dim i As Long
    
    TempName = tn_comp_name & RENAMED_BY_COMPMAN
    Do
        On Error Resume Next
        TempName = tn_wb.VBProject.VBComponents(TempName).Name
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
        Case vbext_ct_StdModule:        TypeString = "Standatd-Module"
    End Select
End Function

