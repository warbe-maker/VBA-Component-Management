VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsServiced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ------------------------------------------------------------------------------------
' Class Module clsServiced: Serviced Workbook properties.
' =========================
'
' Public services:
' - IsHostedCommonCommponent       Returns TRUE when a given component's name is claimed
'                                  being hosted by the serviced Workbook.
' - NoOfItemsTotal           P r/w The total number of items relevant for the service
'
' ------------------------------------------------------------------------------------
Public Comp                 As clsComp

Private dctAllCompsCommon   As New Dictionary
Private dctComps            As New Dictionary
Private dctHosted           As New Dictionary
Private dctPublProcCpys     As New Dictionary
Private dctTypeStrings      As New Dictionary
Private lMaxLenItem         As Long
Private lMaxLenType         As Long
Private lNoOfItemsTotal     As Long
Private sServiceName        As String
Private wbk                 As Workbook

Private Sub Class_Initialize()

    Set Comp = New clsComp
    Set wbk = ActiveWorkbook
    With dctTypeStrings
        .Add vbext_ct_ActiveXDesigner, "ActiveX Designer"
        .Add vbext_ct_ClassModule, "Class Module"
        .Add vbext_ct_Document, "Document Module"
        .Add vbext_ct_MSForm, "UserForm"
        .Add vbext_ct_StdModule, "Standard Module"
    End With
End Sub

Public Property Get Hosted() As Dictionary:                             Set Hosted = dctHosted:                             End Property

Public Property Let HostedCommComps(ByVal h_hosted As String):          Set dctHosted = StringAsDictionary(h_hosted):       End Property

Public Property Get MaxLenItem() As Long

    If lMaxLenItem = 0 Then Serviced.MaxLengths
    MaxLenItem = lMaxLenItem
    
End Property

Public Property Get MaxLenType() As Long

    If lMaxLenType = 0 Then Serviced.MaxLengths
    MaxLenType = lMaxLenType
    
End Property

Public Function TypeString(ByVal t_vbc As VBComponent) As String
' ------------------------------------------------------------------------------
' Returns the component's (t_vbc) type as string.
' ------------------------------------------------------------------------------
    TypeString = dctTypeStrings(t_vbc.Type)
End Function

Friend Property Get NoOfItemsTotal() As Long:               NoOfItemsTotal = lNoOfItemsTotal:                           End Property

Friend Property Let NoOfItemsTotal(ByVal l As Long):        lNoOfItemsTotal = l:                                        End Property

Public Property Get PublicProcCopies() As Dictionary:               Set PublicProcCopies = dctPublProcCpys:             End Property

Public Property Let PublProcCpys(ByVal p_copies As String):         Set dctPublProcCpys = StringAsDictionary(p_copies): End Property

Public Property Get ServiceName() As String:                        ServiceName = sServiceName:                         End Property

Public Property Let ServiceName(ByVal s_service_name As String):    sServiceName = s_service_name:                      End Property

Public Property Get Wrkbk() As Workbook:                            Set Wrkbk = wbk:                                    End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Function CompExists(ByVal e_comp As String) As Boolean
    
    Dim vbc As VBComponent
    If dctComps.Count = 0 Then
        With dctComps
            For Each vbc In wbk.VBProject.VBComponents
                .Add vbc.Name, vbc
            Next vbc
        End With
    End If
    CompExists = dctComps.Exists(e_comp)
    
End Function

Public Function CompsCommon() As Dictionary
' ----------------------------------------------------------------------------
' Returns all components regarded a Common Component.
' ----------------------------------------------------------------------------
    Dim sComp   As String
    Dim v       As Variant
    Dim vbc     As VBComponent
    
    If dctAllCompsCommon.Count = 0 Then
        With dctAllCompsCommon
            '~~ 1. Collect all components known as public
            For Each vbc In wbk.VBProject.VBComponents
                sComp = vbc.Name
                If CommonPublic.All.Exists(sComp) Then
                    .Add sComp, vbNullString
                End If
            Next vbc
            '~~ 2. Add components hosted (in case yet not public)
            For Each v In Hosted
                If Not .Exists(v) Then
                    .Add v, vbNullString
                End If
            Next v
'            '~~ 3. Add components pending by the serviced Workbook - and possibly yet not public
'            For Each v In CommonPending.Components
'                If CompExists(v) And Not .Exists(v) Then
'                    Set Comp = New clsComp
'                    Comp.CompName = v
'                    If Comp.CodePnding.Meets(Comp.CodeCrrent) Then
'                        .Add v, vbNullString
'                    End If
'                    Set Comp = Nothing
'                End If
'            Next v
        End With
    End If
    Set CompsCommon = dctAllCompsCommon
    
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "clsServiced." & s
End Function

Public Function HasBeenRenamedByUpdateService(ByVal i_comp As String) As Boolean
' ------------------------------------------------------------------------------
' Returns True when a component's name (comp_name indicates that it is had been
' renamed by CompMan to enable an update (rename/import) service.
' ------------------------------------------------------------------------------
    HasBeenRenamedByUpdateService = InStr(i_comp, RENAMED_BY_COMPMAN) <> 0
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
    Dim temp    As Variant
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
                temp = arr(j)
                arr(j) = arr(i)
                arr(i) = temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    With dct
        For i = LBound(arr) To UBound(arr)
            vKey = arr(i)
            .Add key:=vKey, Item:=k_dct.Item(vKey)
        Next i
    End With
    
xt: Set k_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub MaxLengths()
' ----------------------------------------------------------------------------
' Returns a Dictionary with all Common Components in the serviced Workbook's
' VB-Project with the component's name as key.
' ----------------------------------------------------------------------------
    Const PROC = "MaxLengths"
    
    On Error GoTo eh
    Dim vbc As VBComponent
    
    If wbk Is Nothing _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The service cannot be used before a serviced Workbook has been provided!"
    
    For Each vbc In wbk.VBProject.VBComponents
        With vbc
            If Not HasBeenRenamedByUpdateService(.Name) Then
                lMaxLenItem = Max(lMaxLenItem, Len(.Name))
                lMaxLenType = Max(lMaxLenType, Len(TypeString(vbc)))
            End If
        End With
    Next vbc
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function StringAsDictionary(ByVal s_items As String, _
                          Optional ByVal s_split As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Transforming the strings within a string (s_items) into a
'            Dictionary by saving the strings as key unifies them. As a
'            compensation of this restriction the number of occurences of a
'            string is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l   As Long
    Dim v   As Variant
    
    Set StringAsDictionary = dct
    If s_items <> vbNullString Then
        If s_split = vbNullString Then s_split = SplitIndctr(s_items)
        With dct
            For Each v In Split(s_items, s_split)
                If Not .Exists(v) Then
                    .Add v, 1
                Else
                    l = dct(v) + 1
                    .Remove v
                    .Add v, l
                End If
            Next v
        End With
        Set StringAsDictionary = KeySort(dct)
    End If
    Set dct = Nothing
    
End Function

