VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dctAll              As Dictionary
Private lMaxLenServicedItem As Long
Private lMaxLenServicedType As Long

Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' Initializes the Class instance with the VB-Components relevant for the
' current service.
' ----------------------------------------------------------------------------
    If Services.CurrentService = mCompManClient.SRVC_UPDATE_OUTDATED Then
        Set dctAll = CollectAllCommComps
    Else
        Set dctAll = CollectAll
    End If
End Sub

Public Property Get All() As Dictionary:            Set All = dctAll:                           End Property

Public Property Get MaxLenServicedItem() As Long:   MaxLenServicedItem = lMaxLenServicedItem:   End Property

Public Property Get MaxLenServicedType() As Long:   MaxLenServicedType = lMaxLenServicedType:   End Property

Private Function CollectAll() As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all VBComponent objects as key and the
' corresponding clsComp objects as item.
' Note: The mDct.DctAdd service orders the items ascending by key (possible
'       because the key object has a Name property.
' ----------------------------------------------------------------------------
    Const PROC = "CollectAll"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim Comp    As clsComp
    Dim wbk     As Workbook
    Dim dct     As New Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    
    With wbk.VBProject
        For Each vbc In .VBComponents
            Set Comp = New clsComp
            With Comp
                .Wrkbk = wbk
                .CompName = vbc.Name
                .VBComp = vbc
                lMaxLenServicedItem = Max(lMaxLenServicedItem, Len(.CompName))
                lMaxLenServicedType = Max(lMaxLenServicedType, Len(.TypeString))
                dct.Add vbc.Name, Comp
            End With
            With Services
                .NoOfItemsTotal = dct.Count
                .DsplyProgress "collected"
            End With
        Next vbc
    End With
        
    If dct.Count > 0 Then mDct.KeySort dct
    
xt: Set CollectAll = dct
    Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Function
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function Changed() As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all components of which the code has been
' modified. The Dictionary has the VBComponent objects as key and the
' corresponding clsComp objects as item. The function uses the All Dictionary
' and removes all items where the components not indicates it has changed.
' ----------------------------------------------------------------------------
    Dim Comp    As clsComp
    Dim v       As Variant
    Dim dct     As Dictionary
    
    Set dct = dctAll
    For Each v In dct
        Set Comp = dct(v)
        If Not Comp.Changed Then
            dct.Remove v
        End If
    Next v
    Set Changed = dct
    
End Function

Private Function CollectAllCommComps() As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all VBComponent objects as key and the
' corresponding clsComp objects as item.
' Note: The mDct.DctAdd service orders the items ascending by key (possible
'       because the key object has a Name property.
' ----------------------------------------------------------------------------
    Const PROC = "CollectAllCommComps"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim Comp    As clsComp
    Dim wbk     As Workbook
    Dim dct     As New Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    
    With wbk.VBProject
        For Each vbc In .VBComponents
            Set Comp = New clsComp
            With Comp
                .Wrkbk = wbk
                .CompName = vbc.Name
                .VBComp = vbc
                Services.ServicedItem = vbc
                lMaxLenServicedItem = Max(lMaxLenServicedItem, Len(.CompName))
                lMaxLenServicedType = Max(lMaxLenServicedType, Len(.TypeString))
                If .KindOfComp = enCommCompHosted Or .KindOfComp = enCommCompUsed Then
                    dct.Add vbc.Name, Comp
                    With Services
                        .NoOfCommonComponents = .NoOfCommonComponents + 1
                        .NoOfItemsTotal = dct.Count
                        .DsplyProgress "collected"
                    End With
                End If
            End With
        Next vbc
    End With
    
    If dct.Count > 0 Then mDct.KeySort dct
    Set CollectAllCommComps = dct

xt: Set dct = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComps" & "." & sProc
End Function

