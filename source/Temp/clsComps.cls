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

Private dctChanged          As Dictionary
Private dctAll              As Dictionary
Private lMaxLenServicedItem As Long
Private lMaxLenServicedType As Long

Private Sub Class_Initialize()
    CollectAll
    Set dctChanged = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set dctChanged = Nothing
End Sub

Public Property Get All() As Dictionary:            Set All = dctAll:                           End Property

Public Property Get MaxLenServicedItem() As Long:   MaxLenServicedItem = lMaxLenServicedItem:   End Property

Public Property Get MaxLenServicedType() As Long:   MaxLenServicedType = lMaxLenServicedType:   End Property

Public Function CollectAll() As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all VBComponent objects as key and the
' corresponding clsComp objects as item.
' Note: The mDct.DctAdd service orders the items ascending by key (possible
'       because the key object has a Name property.
' ----------------------------------------------------------------------------
    Dim vbc     As VBComponent
    Dim Comp    As clsComp
    Dim wbk     As Workbook
    
    Set wbk = Services.Serviced
    If wbk Is Nothing Then
        Set wbk = ActiveWorkbook
    End If
    Set dctAll = New Dictionary
    
    With wbk.VBProject
        For Each vbc In .VBComponents
            Set Comp = New clsComp
            With Comp
                Set .Wrkbk = wbk
                .CompName = vbc.Name
                Set .VBComp = vbc
                lMaxLenServicedItem = Max(lMaxLenServicedItem, Len(.CompName))
                lMaxLenServicedType = Max(lMaxLenServicedType, Len(.TypeString))
            End With
            dctAll.Add vbc.Name, Comp
            With Services
                .NoOfItemsTotal = dctAll.Count
                .DsplyProgress "collected"
            End With
        Next vbc
    End With
    mDct.KeySort dctAll
    
End Function

Private Function Changed(Optional ByRef c_stats As clsStats = Nothing) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary all components of which the code has been modified.
' The Dictionary has the VBComponent objects as key and the corresponding
' clsComp objects as item. The function uses the All Dictionary and removes
' all items where the components not indicates it has changed.
' ----------------------------------------------------------------------------
    Dim Comp    As clsComp
    Dim v       As Variant
    Dim dct     As Dictionary
    
    Set dct = dctAll(c_stats)
    For Each v In dct
        Set Comp = dct(v)
        If Not Comp.Changed Then
            dct.Remove v
        End If
    Next v
    If Not c_stats Is Nothing Then c_stats.Count mCompMan.sic_comps_changed, dct.Count
    Set Changed = dct
    
End Function

Public Function CommonUsed() As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with those VBComponents (clsComp objects) in the
' eerviced Workbook which are indicated Used Common Components.
' ----------------------------------------------------------------------------
    Const PROC = "Outdated"
    
    On Error GoTo eh
    Dim Comp    As clsComp
    Dim dct     As Dictionary
    Dim v       As Variant
    Dim vbc     As VBComponent
    
    Set dct = dctAll
    For Each v In dct
        Set Comp = dct(v)
        With Comp
            Set vbc = .VBComp
            Select Case .KindOfComp
                Case enCommCompUsed
                    With Services
                        .ServicedItem = vbc
                        .NoOfItemsCommonUsed = .NoOfItemsCommonUsed + 1
                        .LogEntry "Used Common Component"
                    End With
                Case enCommCompHosted
                    dct.Remove v
                    With Services
                        .ServicedItem = vbc
                        .NoOfItemsIgnored = .NoOfItemsIgnored + 1
                        .DsplyProgress "Hosted Common Component (ignored)"
                    End With
                Case Else
                    If mCompManDat.RegistrationState(.CompName) = enRegStatePrivate Then
                        dct.Remove v
                        With Services
                            .ServicedItem = vbc
                            .NoOfItemsIgnored = .NoOfItemsIgnored + 1
                            .DsplyProgress "Named like a Common Component but indicated ""private""! (ignored)"
                        End With
                    Else
                        dct.Remove v
                        With Services
                            .ServicedItem = vbc
                            .NoOfItemsIgnored = .NoOfItemsIgnored + 1
                            .DsplyProgress "Internal VBComponent (ignored)"
                        End With
                    End If
            End Select
        End With
    Next v
    
    With Services
        .NoOfItemsOutdated = dct.Count
        .DsplyProgress "collected Common Components used"
    End With
    
xt: Set CommonUsed = dct
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComps" & "." & sProc
End Function

