Attribute VB_Name = "mRef"
Option Explicit

Public Function Exists(ByRef xst_wbk As Workbook, _
                       ByVal xst_ref As Variant, _
              Optional ByRef xst_ref_result As Reference = Nothing) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Reference (xst_ref) - which might be a Reference object
' or a string - exists in the VB-Project of the Workbook (xst_wbk). When a string
' is provided the reference exists when the string is equal to the Name argument
' or it is LIKE the Description argument of any Reference. The existing
' Refwerence is returned as object (xst_ref_result).
' ------------------------------------------------------------------------------
    Dim ref As Reference
    
    For Each ref In xst_wbk.VBProject.References
        If TypeName(xst_ref) = "Reference" Then
            Exists = ref.Name = xst_ref.Name
        ElseIf TypeName(xst_ref) = "String" Then
            Exists = ref.Name = xst_ref Or ref.Description Like xst_ref
        End If
        If Exists Then
            Set xst_ref_result = ref
            Exit Function
        End If
    Next ref

End Function

