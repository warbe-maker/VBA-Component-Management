Attribute VB_Name = "mVBP"
Option Explicit
Option Private Module
Option Compare Text
' -----------------------------------------------------------------------------------
' Standard  Module mVBP Existence checks for VB-Project objects.
'
' Methods:
' - ComponentExists     Returns TRUE when the object exists
' - CustomViewExists    Returns TRUE when the object exists
' - ProcedureExists     Returns TRUE when the object exists
' - ReferenceExists     Returns TRUE when the object exists
'
' Uses:     Standard Module mErH
'
' Requires: Reference to "Microsoft Scripting Runtine"
'           Reference to "Microsoft Visual Basic for Applications Extensibility ..."
'
' W. Rauschenberger, Berlin August 2019
' -----------------------------------------------------------------------------------

Public Function ProcedureExists(ByVal v As Variant, _
                                ByVal sName As String) As Boolean
' ---------------------------------------------------------------
' Returns TRUE when the Procedure named (sName) exists in the
' CodeModule (vbcm).
' ---------------------------------------------------------------
Const PROC      As String = "ProcedureExists"       ' This procedure's name for the error handling and execution tracking
Dim vbcm        As CodeModule
Dim iLine       As Long             ' For the existence check of a VBA procedure in a CodeModule
Dim sLine       As String           ' For the existence check of a VBA procedure in a CodeModule
Dim vbProcKind  As vbext_ProcKind   ' For the existence check of a VBA procedure in a CodeModule

    On Error GoTo eh
    BoP ErrSrc(PROC)
    ProcedureExists = False

    If Not TypeName(v) = "Nothing" Then
        If TypeOf v Is VBComponent Then
            Set vbcm = v.CodeModule
            With vbcm
                For iLine = 1 To .CountOfLines
                    If .ProcOfLine(iLine, vbProcKind) = sName Then
                        ProcedureExists = True
                        GoTo xt
                    End If
                Next iLine
                GoTo xt
            End With
        ElseIf TypeOf v Is CodeModule Then
            Set vbcm = v
            With vbcm
                For iLine = 1 To .CountOfLines
                    If .ProcOfLine(iLine, vbProcKind) = sName Then
                        ProcedureExists = True
                        GoTo xt
                    End If
                Next iLine
                GoTo xt
            End With
        End If
    End If
    Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The item (parameter v) for the Procedure's existence check is neither a Component object nor a CodeModule object!"

xt: EoP ErrSrc(PROC)
    Exit Function

eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Function CodeModuleIsEmpty(ByVal v As Variant) As Boolean
' --------------------------------------------------------------
' Returns TRUE when the CodeModule (v) has only one line with a
' lenght of n
' --------------------------------------------------------------
Const PROC  As String = "CodeModuleIsEmpty"
Dim vbc     As VBComponent
Dim vbcm    As CodeModule

    On Error GoTo eh

    If Not VarType(v) = vbObject _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The parameter is not an object (VBComponwent or CodeModule)!"
    If Not TypeOf v Is VBComponent And Not TypeOf v Is CodeModule _
    Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "The parameter (v) is neither a VBComponent nor a CodeModule object!"
    
    If TypeOf v Is CodeModule Then
        Set vbcm = v
    ElseIf TypeOf v Is VBComponent Then
        Set vbc = v
        Set vbcm = vbc.CodeModule
    End If
    
    With vbcm
        If .CountOfLines = 0 Then
            CodeModuleIsEmpty = True
        ElseIf .CountOfLines = 1 And Len(.Lines(1, 1)) < 2 Then
            CodeModuleIsEmpty = True
        End If
    End With

xt: Exit Function
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Sub CodeModuleTrim(ByVal v As Variant, _
                          ByVal wb As Workbook)
' -------------------------------------------------------
' Remove any leading or trailing empty code lines from
' the codemodule (v) - which may be a VBComponent, a
' CodeModule or a string (= a Component's name).
' When (v) is a string and no Workbook (vWb) is provided,
' the Workbook defaults to the ActiveWorkbook. Else,
' the Workbook may be provided as an open Workbook's name
' or a Workbook object.
' -------------------------------------------------------
Const PROC  As String = "CodeModuleTrim"
Dim vbc     As VBComponent
Dim vbcm    As CodeModule

    On Error GoTo eh
   
    Select Case TypeName(v)
        Case "String"
            If v = vbNullString _
            Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "A CodeModule (v) is not provided!"
            
            '~~ The existence check returns the VBComponent object when it exists
            If Not mVBP.ComponentExists(wb, v, vbc) _
            Then Err.Raise mErH.AppErr(5), ErrSrc(PROC), "The CodeModule '" & v & "' (v) does not exist in the Workbook '" & wb.Name & "'!"
            Set vbcm = vbc.CodeModule

        Case "VBComponent"
            Set vbcm = v.CodeModule
        Case "CodeModule"
            Set vbcm = v
    End Select
    
    
    With vbcm
        If wb Is ThisWorkbook Then
            If Len(.Lines(1, 1)) = 0 Then
                MsgBox "The CodeModule of '" & vbcm.Parent.Name & "' has an empty code line at the top " & _
                       "which cannot be removed since the Workbook is ME! (" & wb.Name & ")." & vbLf & _
                       "Since the check whether the code has changed or not is done by comparing the code " & _
                       "with its ExportFile (which is done by transferring both into anarray). " & _
                       "This comparison may indicate a code change though the relevant code has not changed.", _
                       vbCritical, _
                       "Empty code line cannot be removed"
            End If
            GoTo xt ' May cause Excel to crash !
        End If
        
        '~~ Remove any leading empty code line
        Do While Len(.Lines(1, 1)) = 0
            .DeleteLines 1, 1
        Loop
        '~~ Remove any leading empty code line
        Do While Len(.Lines(.CountOfLines, 1)) = 0
            .DeleteLines .CountOfLines, 1
        Loop
    End With
    
xt: Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Function ComponentExists(ByVal vWb As Variant, _
                                ByVal vComp As Variant, _
                       Optional ByRef vbcResult As VBComponent) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE and the Component object (cmpResult) when the Component
' (vComp) - which may be a Component object or a Component's name - exists
' in the Workbook (vWb) - which may be a Workbook object or a Workbook's
' name or fullname of an open Workbook.
' ------------------------------------------------------------------------
Const PROC  As String = "ComponentExists"       ' This procedure's name for the error handling and execution tracking
Dim wb      As Workbook
Dim sTest   As String
Dim sName   As String
Dim vbc     As VBComponent

    On Error GoTo eh
    ComponentExists = False
        
    Select Case TypeName(vWb)
        Case "Workbook"
            Set wb = vWb
        Case "String"
            If Not mWrkbk.IsOpen(vWb, wb) _
            Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The provided Workbook (vWb) is not open!"
        Case Else
            Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The Workbook (vWb) is neither an object nor a string!"
    End Select
    
    Select Case TypeName(vComp)
        Case "VBComponent"
            sName = vComp.Name
        Case "String"
            sName = vComp
        Case Else
            Err.Raise mErH.AppErr(3), ErrSrc(PROC), "The Component (vComp) is neither an object nor a string!"
    End Select
    
    On Error Resume Next
    sName = wb.VBProject.VBComponents(sName).Name
    If Err.Number = 0 Then
        ComponentExists = True
        Set vbcResult = wb.VBProject.VBComponents(sName)
    End If

xt: Exit Function

eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Function CustomViewExists(ByVal vWb As Variant, _
                                 ByVal vCv As Variant) As Boolean
' ----------------------------------------------------------------
' Returns TRUE when the CustomView (vCv) - which may be a CustomView
' object or a CustoView's name - exists in the Workbook (wb). If
' vCv is provided as CustomView object, only its name is used for
' the existence check in Workbook (wb).
' ---------------------------------------------------------------
Const PROC  As String = "CustomViewExists"      ' This procedure's name for the error handling and execution tracking
Dim wb      As Workbook
Dim sTest   As String

    On Error GoTo eh
    BoP ErrSrc(PROC)
    CustomViewExists = False
    
    If Not TypeName(vWb) = "Workbook" And VarType(vWb) <> vbString Then
        Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The provided Workbook paramenter (vWB) is neither a Workbook object nor a string!"
    End If
    
    If TypeName(vWb) = "Workbook" Then
        Set wb = vWb
    Else
        Set wb = mWrkbk.GetOpen(vWb)   ' raises an error when not open
    End If
    
    If TypeName(vCv) = "Nothing" Then
        Err.Raise mErH.AppErr(2), ErrSrc(PROC), "The paramenter (vCv) is not provided!"
    End If
    
        If TypeOf vCv Is CustomView Then
            On Error Resume Next
            sTest = vCv.Name
            CustomViewExists = Err.Number = 0
            GoTo xt
        End If
    If VarType(vCv) = vbString Then
        On Error Resume Next
        sTest = wb.CustomViews(vCv).Name
        CustomViewExists = Err.Number = 0
        GoTo xt
    End If
    Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The CustomView (parameter vCv) for the CustomView's existence check is neither a string (CustomView's name) nor a CustomView object!"
        
xt: EoP ErrSrc(PROC)
    Exit Function
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Function ReferenceExists(ByVal vWb As Variant, _
                                ByVal vRef As Variant, _
                       Optional ByRef refResult As Reference) As Boolean
' ----------------------------------------------------------------------
' Returns TRUE when the Reference (vRef) - which may be a Reference
' object or a Refernece's GUID - exists in the VBProject of the Workbook
' (vWb) - which may be a Workbook object or a Workbook's name or
' fullname. When vRef is provided as object only its GUID is used for
' the existence check in Workbook (vWb).
' ----------------------------------------------------------------------
Const PROC  As String = "ReferenceExists"   ' This procedure's name for the error handling and execution tracking
Dim ref     As Reference
Dim refTest As Reference
Dim wb      As Workbook

    On Error GoTo eh
    BoP ErrSrc(PROC)
    ReferenceExists = False
    
    '~~ Assert vWb
    If Not TypeName(vWb) = "Workbook" And Not TypeName(vWb) = "String" _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The provided Workbook (vWb) is neither an object nor a string!"
    If TypeName(vWb) = "String" Then
        If Not mWrkbk.IsOpen(vWb, wb) _
        Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "The provided Workbook (vWb) is not open!"
    Else
        Set wb = vWb
    End If

    If TypeName(vRef) = "Nothing" _
    Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "The Reference (parameter vRef) for the Reference's existence check is ""Nothing""!"
    
    If Not TypeOf vRef Is Reference And VarType(vRef) <> vbString _
    Then Err.Raise mErH.AppErr(3), ErrSrc(PROC), "The Reference (parameter vRef) for the Reference's existence check is neither a valid GUID (a string enclosed in { } ) nor a Reference object!"
    
    If VarType(vRef) = vbString Then
        If left$(vRef, 1) <> "{" Or Right$(vRef, 1) <> "}" _
        Then Err.Raise mErH.AppErr(4), ErrSrc(PROC), "The Reference (parameter vRef) for the Reference's existence check is a string but not syntactically correct GUID (a string enclosed in { } )!"
    End If
    
    If TypeOf vRef Is Reference Then
        Set refTest = vRef
        For Each ref In wb.VBProject.References
            If ref.GUID = refTest.GUID Then
                ReferenceExists = True
                Set refResult = ref
                GoTo xt
            End If
        Next ref
    ElseIf VarType(vRef) = vbString Then
        For Each ref In wb.VBProject.References
            If ref.GUID = vRef Then
                ReferenceExists = True
                Set refResult = ref
                GoTo xt
            End If
        Next ref
    End If

xt: EoP ErrSrc(PROC)
    Exit Function

eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & ">mVBP" & ">" & sProc
End Function
