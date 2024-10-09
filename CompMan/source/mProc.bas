Attribute VB_Name = "mProc"
Option Explicit ' Under  c o n s t r u c t i o n !
' ----------------------------------------------------------------------------
' Standard Module mProc: Code copy dedection and management service.
' ======================
' ----------------------------------------------------------------------------
Private Const BttnDsplyCodeDiff As String = "Display the code difference"
Private Const Bttn2             As String = "Button 2"
Private Const Bttn3             As String = "Button 3"

Public Const PROC_COPY_INDICATOR = "' Origin: Common Component ""<comp-name>"" =========================="

Private Sub Collect(ByRef c_dct As Dictionary, _
                    ByVal c_key As String, _
                    ByVal c_comp As String, _
                    ByVal c_proc As String, _
                    ByVal c_code As String, _
                    ByVal c_kind As String, _
                    ByVal c_scope As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    
    Dim Procedure   As clsProc
    
    If Not c_dct.Exists(c_key) Then
        Set Procedure = New clsProc
        With Procedure
            .CompName = c_comp
            .ProcName = c_proc
            .Code = c_code
            .Scope = c_scope
            .KindOfProc = c_kind
        End With
        c_dct.Add c_key, Procedure
        Set Procedure = Nothing
    Else
        '~~ This (must be) the second procedure of a Property
        Set Procedure = c_dct(c_key)
        With Procedure
            .Code = c_code ' is added to the already saved code lines
            If c_scope = "Public" Then .Scope = c_scope
            c_dct.Remove c_key
            If .Scope = "Public" Then
                c_dct.Add c_key, Procedure
                Set Procedure = Nothing
            End If
        End With
    End If

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mProc." & sProc
End Function

Private Function KeySort(ByRef k_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (k_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim Temp    As Variant
    Dim i       As Long
    Dim j       As Long
    
    If k_dct Is Nothing Then GoTo xt
    If k_dct.Count = 0 Then GoTo xt
    
    With k_dct
        ReDim arr(0 To .Count - 1)
        For i = 0 To .Count - 1
            arr(i) = .Keys(i)
        Next i
    End With
    
    '~~ Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add Key:=vKey, Item:=k_dct.Item(vKey)
    Next i
    
xt: Set k_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Kind(ByVal k_kind As vbext_ProcKind) As String

    Select Case k_kind
        Case vbext_pk_Get:  Kind = "Property-Get"
        Case vbext_pk_Let:  Kind = "Property-Let"
        Case vbext_pk_Proc: Kind = "Proc"
        Case vbext_pk_Set:  Kind = "Property-Set"
    End Select

End Function

Private Sub ProcSpecs(ByVal p_code As String, _
                      ByVal p_name As String, _
                      ByRef p_kind As String, _
                      ByRef p_scope As String)
' ----------------------------------------------------------------------------
' Returns the specifics of the current procedure, kind (p_kind) and scope
' (p_scope), derived from the respective line in the code (p_code).
' ----------------------------------------------------------------------------
    Const PROC = "KindOfProcAsString"
    
    Dim i       As Long
    Dim aCode   As Variant
    Dim aLine   As Variant
    
    If p_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The procedure's name is yet unknown!"
    
    aCode = Split(p_code, vbCrLf)
    For i = LBound(aCode) To UBound(aCode)
        If InStr(aCode(i), p_name & "(") <> 0 Then
            aLine = Split(aCode(i), " ")
            If aLine(LBound(aLine)) = "Public" Or _
               aLine(LBound(aLine)) = "Private" Then
               p_scope = aLine(LBound(aLine))
               p_kind = aLine(LBound(aLine) + 1)
            Else
                p_scope = "Public"
                p_kind = aLine(LBound(aLine))
            End If
            Exit For
        End If
    Next i

End Sub

Private Function ProcCopies(Optional ByVal p_wbk As Workbook = Nothing, _
                            Optional ByVal p_scope As String = "Public", _
                            Optional ByVal p_common As Boolean = True, _
                            Optional ByRef p_public_unique As Dictionary = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all Procedures which do have at least one equally
' named procedure in another component.
' ----------------------------------------------------------------------------
    Const PROC = "ProcCopies"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim sProc       As String
    Dim v           As Variant
    Dim lCpys       As Long
    Dim dctUnique   As New Dictionary
    
    '~~ Obtain all procedures if the requested scope in Common Components (p_common = True)
    '~~ or non-common components (p_common = False)
    Set dct = mProc.ProcsCollect(p_wbk, p_scope, p_common, p_public_unique)
    
    '~~ Isolate those with copies
    For Each v In dct
        sProc = Split(v, ".")(0)
        If Not dctUnique.Exists(sProc) Then
            lCpys = 1
            dctUnique.Add sProc, 1
        Else
            lCpys = dctUnique.Item(sProc)
            lCpys = lCpys + 1
            dctUnique.Remove sProc
            dctUnique.Add sProc, lCpys
        End If
    Next v
    
    For Each v In dctUnique
        If dctUnique.Item(v) = 1 Then dctUnique.Remove v
    Next v
    
'    '~~ Remove single (non-copy) procedures
'    If p_scope = "Private" Then
'        For Each v In dct
'            sProc = Split(v, ".")(0)
'            If Not dctUnique.Exists(sProc) Then
'                dct.Remove v
'            End If
'        Next v
'    End If
    
    Set ProcCopies = dct
    Set dct = Nothing
    
    Set p_public_unique = dctUnique
    Set dctUnique = Nothing
    
xt: Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub PublicProcCopyManagement(Optional ByVal c_wbk As Workbook = Nothing)
' ----------------------------------------------------------------------------
' Checks for each declared/provided copy of a public procedure in a public
' Common Component whether it is up-to-date. Displays a dialog if not, with
' the options to display the difference, update, or skip.
' ----------------------------------------------------------------------------
    Const PROC = "PublicProcCopyManagement"
    
    On Error GoTo eh
    Dim v As Variant
    Dim sCpyProc    As String ' Name of a copied procedure
    Dim sCpyComp    As String ' Name of the component it is used
    Dim sCommComp   As String ' Origin Common Component
    Dim ProcCopy    As New clsProc
    Dim ProcPublic  As New clsProc
    Dim vbcm        As CodeModule
    
    If c_wbk Is Nothing Then Set c_wbk = ActiveWorkbook
    
    For Each v In Serviced.PublicProcCopies
        sCommComp = Split(v, ":")(1)
        sCpyProc = Split(Split(v, ":")(0), ".")(1)
        sCpyComp = Split(Split(v, ":")(0), ".")(0)
        
        If Not mFact.PublicCommonComponent(sCommComp) _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided origin Common Component """ & sCommComp & """ is not a known public Common Component!"
        If Not mFact.ProcInComp(c_wbk, sCpyComp, sCpyProc, vbcm) _
        Then Err.Raise AppErr(2), ErrSrc(PROC), "The procedure """ & sCpyProc & """ is not known in "
        With ProcPublic
            .CompName = sCommComp
            .ProcName = sCpyProc
            .Source = CommonPublic.LastModExpFile(sCommComp)
        End With
        With ProcCopy
            .CompName = sCpyComp
            .ProcName = sCpyProc
            .Source = vbcm
            If .DiffersFrom(ProcPublic) Then
                Do
                    Select Case DsplyDiffOptions(sCpyProc, sCommComp)
                        Case BttnDsplyCodeDiff
                            .DsplyDiffs "ProcCopy", "The copied procedure's current code", ProcPublic, "ProcPublic", "The code of the current public Common Component's Procedure"
                        Case Else
                            Exit Do
                    End Select
                Loop
            End If
        End With
    Next v

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function DsplyDiffOptions(ByVal d_proc As String, _
                                  ByVal d_comp As String) As Variant
    
    Dim Msg As udtMsg
    
    With Msg
        .Section(1).Text.Text = "The copy of the Public Procedure """ & d_proc & """, registered for being monitored, " & _
                                "differs from its origin in the Common Component """ & d_comp & """!"
        With .Section(2)
            .Label.FontColor = rgbBlue
            .Label.Text = BttnDsplyCodeDiff
            .Text.Text = "Display the code difference before making a decision"
        End With
    End With
    
    DsplyDiffOptions = mMsg.Dsply(dsply_title:="Code of Public Procedure copy """ & d_proc & """ outdated!" _
                                , dsply_msg:=Msg _
                                , dsply_buttons:=mMsg.Buttons(BttnDsplyCodeDiff, Bttn2, Bttn3))
                                
End Function

Public Sub PublicProcCopyManagementTest()
    Dim wbk As Workbook
    
    Set wbk = ActiveWorkbook
    
    mCompMan.ServiceInitiate wbk, "PublicProcCopyManagement", "", "clsTestAid.Arry:mBasic, clsTestAid.FileAsArray:mVarTrans"
    PublicProcCopyManagement wbk
    
End Sub

Public Sub CopyDedectionAndManagement(Optional ByVal c_wbk As Workbook = Nothing)
' ----------------------------------------------------------------------------
' Dedection of "potential" copies of Public procedures in Common Components
' and "potential" Private copies of the Public ones.
'
' Background:
' A major aim for Common Components is to keep then autonomous in the sense
' that they do not need/use any other component. To meet this they will use
' copies of procedures provided Public in other Common Components. However,
' these copies are supposed to remain identicall with their origin. Any
' modification made in the origin or in the Public or Private copy should
' be recognized in order to eliminate the "inconsitency".
' ----------------------------------------------------------------------------
    Const PROC = "CopyDedectionAndManagement"
    
    On Error GoTo eh
    
    Dim dctPrivate  As Dictionary
    Dim dctPublic   As Dictionary
    Dim dctUnique   As New Dictionary
    Dim v           As Variant
    
    If c_wbk Is Nothing Then Set c_wbk = ActiveWorkbook
    mCompMan.ServiceInitiate c_wbk, PROC
      
    '~~ Get all Public procedures in Common Components which are used more than once
    Set dctPublic = mProc.ProcCopies(c_wbk, "Public", True, dctUnique)
    For Each v In dctUnique
        Debug.Print ErrSrc(PROC) & ": " & Format(dctUnique.Item(v), "00") & " Public copies found in Common Components of procedure: " & v
    Next v
    
    '~~ Get all Private procedures which are potential copies of a Public procedure in a Common Component
    Set dctPrivate = mProc.ProcCopies(c_wbk, "Private", False, dctUnique)
    For Each v In dctUnique
        Debug.Print ErrSrc(PROC) & ": " & Format(dctUnique.Item(v), "00") & " Private copies of Public procedures in Common Components: " & v
    Next v
    
    For Each v In dctUnique
    
    Next v
    
    Set dctPublic = Nothing
    Set dctPrivate = Nothing
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function ProcsCollect(Optional ByVal p_wbk As Workbook = Nothing, _
                             Optional ByVal p_scope As String = "Public", _
                             Optional ByVal p_common As Boolean = True, _
                             Optional ByVal p_public_unique As Dictionary = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all Procedures in a Workbook (p_wbk) with the
' provided scope (p_scope) either within common (p_common = True) or non-
' common (p_common = False) components.
' When non-common procedures are requested only those identified public
' beforehand are relevant.
' ----------------------------------------------------------------------------

    Dim dct         As New Dictionary
    Dim sKind       As String
    Dim sScope      As String
    Dim sCompName   As String
    Dim sProcName   As String
    Dim vbc         As VBComponent
    Dim vbcm        As CodeModule
    Dim sCode       As String
    Dim sKey        As String
    Dim lKind       As vbext_ProcKind
    Dim dctCommComps    As New Dictionary
    Dim lFrst           As Long
    Dim lLines          As Long
    
    If p_wbk Is Nothing Then Set p_wbk = ActiveWorkbook
    If p_common Then Set dctCommComps = CommonServiced.Components
    
    For Each vbc In p_wbk.VBProject.VBComponents
        sCompName = vbc.Name
        If (p_common And dctCommComps.Exists(sCompName)) _
        Or p_common = False Then
            '~~ Either Common Coponent (when only Common Components are requested)
            '~~ or "Private" with a same named Common counterpart
            Set vbcm = vbc.CodeModule
            With vbcm
                lFrst = .CountOfDeclarationLines + 1
                Do While lFrst < .CountOfLines
                    '~~ First line of a procedure
                    sProcName = .ProcOfLine(lFrst, lKind)
                    lLines = .ProcCountLines(sProcName, lKind)
                    sCode = .Lines(lFrst, lLines)
                    ProcSpecs sCode, sProcName, sKind, sScope
                    sKey = sProcName & "." & sCompName
                    Select Case True
                        Case Not p_common And p_public_unique.Exists(sProcName)
                            Collect dct, sKey, sCompName, sProcName, sCode, sKind, sScope
                        Case p_common And sScope = "Public"
                            Select Case True
                                Case sScope = p_scope Or sKind = "Property"
                                    sKey = sProcName & "." & sCompName
                                    Collect dct, sKey, sCompName, sProcName, sCode, sKind, sScope
                                Case sScope = "Private" And sKind = "Property"
                                    sKey = sProcName & "." & sCompName
                                    Collect dct, sKey, sCompName, sProcName, sCode, sKind, sScope
                            End Select
                    End Select
                    lFrst = lFrst + lLines
                Loop
            End With
        End If ' Is a Common Component of the serviced Workbook
    Next vbc
    
    Set ProcsCollect = KeySort(dct)
    Set dct = Nothing
    
End Function

Public Sub TestProcClass()

    Dim Prcdr   As New clsProc
    Dim sComp   As String
    Dim sProc   As String
    Dim wbk     As Workbook
    
    sComp = "clsTestAid"
    sProc = "Arry"
    
    Set wbk = ActiveWorkbook

    With Prcdr
        .CompName = sComp
        .ProcName = sProc
        .Source = wbk.VBProject.VBComponents(sComp).CodeModule
        Debug.Print ErrSrc(Prcdr) & ": " & "=======================" & vbLf & .Code & vbLf & "====================="
    End With
    Set Prcdr = Nothing
    
End Sub
