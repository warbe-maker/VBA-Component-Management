Attribute VB_Name = "mComp"
Option Explicit
' ------------------------------------------------------------------------
' Standard-Module mComp: Services for (not only Common) VBComponents.
'
' Public services:
' - CommCompRegStateEnum
'   Returns a provided CommCompRegStateString as enumerated value
' - CommCompRegStateString
'   Returns a provided enumerated CommCompRegStateString as string
' - Exists
'   Returns TRUE when a components name exists in a Workbook's VBProject
' - IsSheetDocMod
'   Returns TRUE when a provided VBComponent object is of a type
'   Document-Module and represents a Worksheet
' - IsWrkbDocMod
'   Returns TRUE when a provided VBComponent object is of a type
'   Document-Module and represents the Workbook
'
' - ManageUsedCommonComponent
'
' - TempName
'   Returns a temporary name for a provided VBComponent's name which
'   is not already used in the provided Workbook/VBProject
' - TypeString
'   Returns a provided VBComponent object's type as string
'
' ------------------------------------------------------------------------
Public Const RENAMED_BY_COMPMAN = "_RnmdByCompMan"

Public Enum enCommCompRegState
    enRegStateHosted = 1    ' The Common Component is "hosted"
    enRegStateUsed          ' The Common Component is "used"
    enRegStatePrivate       ' Not a Common Component (though the name matches)
End Enum

Public Function Exists(ByVal ex_comp As Variant, _
                       ByVal ex_wbk As Workbook, _
              Optional ByRef ex_vbc As VBComponent) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE and the VBComponent (ex_vbc) when the provided VB-
' Component (ex_comp) exists in the Workbook's (ex_wbk) VB-Project.
' ------------------------------------------------------------------------------
    On Error Resume Next
    Select Case True
        Case TypeOf ex_comp Is VBComponent:  Set ex_vbc = ex_wbk.VBProject.VBComponents(ex_comp.Name)
        Case VarType(ex_comp) = vbString:    Set ex_vbc = ex_wbk.VBProject.VBComponents(ex_comp)
    End Select
    Exists = Err.Number = 0
End Function

Public Function IsSheetDocMod(ByVal i_vbc As VBComponent, _
                              ByVal i_wbk As Workbook, _
                     Optional ByRef i_wsh As Worksheet) As Boolean
' ------------------------------------------------------------------------------
' When the VBComponent (vbc) represents a Worksheet the function returns TRUE
' and the corresponding Worksheet (i_wsh).
' ------------------------------------------------------------------------------
    Dim wsh As Worksheet

    IsSheetDocMod = i_vbc.Type = vbext_ct_Document And i_vbc.Name <> i_wbk.CodeName
    If IsSheetDocMod Then
        Debug.Print "i_vbc.Name: " & i_vbc.Name
        For Each wsh In i_wbk.Worksheets
            Debug.Print "wsh.CodeName: " & wsh.CodeName
            If wsh.CodeName = i_vbc.Name Then
                Set i_wsh = wsh
                Exit For
            End If
        Next wsh
    End If

End Function

Public Function TempName(ByVal tn_wbk As Workbook, _
                         ByVal tn_vbc_name As String) As String
' ----------------------------------------------------------------------------
' Returns a yet not existing temporary name for a component (tn_t_vbc_name).
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

Public Function TypeString(ByVal t_vbc As VBComponent, _
                           ByVal t_wbk As Workbook) As String
    Select Case t_vbc.Type
        Case vbext_ct_ActiveXDesigner:  TypeString = "ActiveX-Designer"
        Case vbext_ct_ClassModule:      TypeString = "Class-Module"
        Case vbext_ct_Document:         If IsSheetDocMod(t_vbc, t_wbk) _
                                        Then TypeString = "Document-Module (Worksheet)" _
                                        Else TypeString = "Document-Module (Workbook)"
        Case vbext_ct_MSForm:           TypeString = "UserForm"
        Case vbext_ct_StdModule:        TypeString = "Standard-Module"
    End Select
End Function

