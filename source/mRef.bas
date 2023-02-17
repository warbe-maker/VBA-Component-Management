Attribute VB_Name = "mRef"
Option Explicit

Public Function Exists(ByVal x_this_ref As Variant, _
                       ByVal x_in_this_wbk As Workbook, _
              Optional ByRef x_ref_result As Reference = Nothing) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Reference (x_this_ref) - which might be a Reference object
' or a string - exists in the VB-Project of the Workbook (x_in_this_wbk). When a string
' is provided the reference exists when the string is equal to the Name argument
' or it is LIKE the Description argument of any Reference. The existing
' Refwerence is returned as object (x_ref_result).
' ------------------------------------------------------------------------------
    Dim ref As Reference
    
    For Each ref In x_in_this_wbk.VBProject.References
        If TypeName(x_this_ref) = "Reference" Then
            Exists = ref.Name = x_this_ref.Name
        ElseIf TypeName(x_this_ref) = "String" Then
            Exists = ref.Name = x_this_ref Or ref.Description Like x_this_ref
        End If
        If Exists Then
            Set x_ref_result = ref
            Exit Function
        End If
    Next ref

End Function

