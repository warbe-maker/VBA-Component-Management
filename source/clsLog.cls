VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sService                As String
Private sFile                   As String
Private sFolder                 As String
Private bNewLog                 As Boolean
Private sEntry                  As String
Private sServicedItem           As String
Private sServicedItemType       As String
Private sServicedItemName       As String
Private lMaxLenServicedItemType As Long
Private lMaxLenServicedItemName As Long

Private Sub Class_Initialize()
    bNewLog = True
End Sub

Friend Property Let Entry(ByVal s As String)
' ------------------------------------------
' Append an entry to the log file.
' ------------------------------------------
    Static sItem    As String
    Dim sCompName   As String
    
    If ServicedItem = vbNullString Then
        sEntry = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & s
    Else
        If ServicedItem = sItem Then
            sEntry = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.Space$(Len(sItem)) & s
        Else
            sEntry = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & ServicedItem & s
        End If
    End If
    mFile.Txt(ft_file:=sFile, ft_append:=True) = sEntry
    sItem = sServicedItem
    sCompName = Split(ServicedItemName, "(")(0)
    mTrc.TraceLogInfo = sCompName & ": " & s
    
End Property

Friend Property Get File() As String:               File = sFile:               End Property

Friend Property Let File(ByVal s As String):        sFile = s:                  End Property

Friend Property Get Folder() As String:             Folder = sFolder:           End Property

Friend Property Let Folder(ByVal s As String):      sFolder = s:                End Property

Friend Property Get LogFile() As File
    Dim sLogFullName    As String
    If InStr(Me.File, ":") <> 0 Then
        sLogFullName = Me.File
    Else
        sLogFullName = Me.Folder & "\" & Me.File
    End If
    With New FileSystemObject
        If Not .FileExists(sLogFullName) Then .CreateTextFile sLogFullName
        Set LogFile = .GetFile(sLogFullName)
    End With
End Property

Friend Property Get MaxLenServicedItemName() As Long
    MaxLenServicedItemName = lMaxLenServicedItemName
End Property

Friend Property Get MaxLenServicedItemType() As Long
    MaxLenServicedItemType = lMaxLenServicedItemType
End Property

Friend Property Get Service( _
             Optional ByVal new_log As Boolean = False) As String
    Service = sService
End Property

Friend Property Let Service( _
             Optional ByVal new_log As Boolean = False, _
                      ByVal sv_name As String)
' -------------------------------------------------------
'
' -------------------------------------------------------
    Const PROC = "Service-Let"
    
    On Error GoTo eh
    
    sService = "'" & sv_name & "(" & ThisWorkbook.Name & ")' for '" & mService.Serviced.Name & "': "
    
    bNewLog = new_log
    Me.Folder = mService.Serviced.Path
    
    If sFile = vbNullString Then
        '~~ In case not a specific file had explicitely assigned before
        sFile = Me.Folder & "\" & mService.SERVICES_LOG_FILE
        If InStr(sFile, "\Addin") <> 0 Then Stop
    End If
    With New FileSystemObject
        If bNewLog Then
            If .FileExists(sFile) Then .DeleteFile sFile
            .CreateTextFile (sFile)
            bNewLog = False
        End If
    End With
    
    If mFile.Txt(ft_file:=sFile) <> vbNullString _
    Then mFile.Txt(ft_file:=sFile _
                 , ft_append:=True _
                  ) = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.String$(Len(sService), "=")
    
    mFile.Txt(ft_file:=sFile _
            , ft_append:=True _
             ) = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & sService & vbCrLf & _
                 Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.String$(Len(sService), "-")

xt: Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get ServicedItem() As Variant
    If sServicedItemType = vbNullString And sServicedItemName = vbNullString Then
        ServicedItem = sServicedItem
    Else
        sServicedItem = mBasic.Align(sServicedItemType, lMaxLenServicedItemType) & " " & mBasic.Align(sServicedItemName, lMaxLenServicedItemName, , " ", ".") & ": "
        ServicedItem = sServicedItem
    End If
End Property

Friend Property Let ServicedItem(ByRef v As Variant)
' ----------------------------------------------------
' Determine the serviced item's type and name string
' for logging and compute the maximum length for both.
' ----------------------------------------------------
    
    Dim vbc As VBComponent
    Dim nm  As Name
    Dim ws  As Worksheet
    Dim ref As Reference
    Dim shp As Shape
    Dim oob As OLEObject
    Dim tbt As ToggleButton
    Dim tbx As TextBox
    Dim sbt As SpinButton
    Dim scb As ScrollBar
    Dim obt As OptionButton
    Dim lbx As ListBox
    Dim lbl As Label
    Dim img As Image
    Dim cbt As CommandButton
    Dim rng As Range
    
    Select Case TypeName(v)
        Case "VBComponent"
            Set vbc = v
            Select Case vbc.Type
                Case vbext_ct_ActiveXDesigner:  ServicedItemType = "ActiveX-Designer":  ServicedItemName = vbc.Name
                Case vbext_ct_ClassModule:      ServicedItemType = "Class-Module":      ServicedItemName = vbc.Name
                Case vbext_ct_MSForm:           ServicedItemType = "UserForm":          ServicedItemName = vbc.Name
                Case vbext_ct_StdModule:        ServicedItemType = "Standard-Module":                                           ServicedItemName = vbc.Name
                Case vbext_ct_Document
                    If mComp.IsSheetDocMod(vbc) Then
                                                ServicedItemType = "Worksheet CodeName":                                        ServicedItemName = vbc.Name
                    Else
                                                ServicedItemType = "Workbook":                                                  ServicedItemName = vbc.Name
                    End If
            End Select
        Case "Name"
            Set nm = v
            ServicedItemType = "Name"
            ServicedItemName = Replace(nm.Name, nm.Parent.Name & "!", vbNullString) & "(" & nm.RefersTo & ")"
            If MaxLenServicedItemName = 60 Then Stop

        Case "OLEObject":       Set oob = v:    ServicedItemType = "(Sheet)ActiveX":    ServicedItemName = "(" & oob.Parent.Name & ")" & oob.Name
        Case "Reference":       Set ref = v:    ServicedItemType = TypeName(ref):                                               ServicedItemName = ref.Name
        Case "Shape":           Set shp = v:    ServicedItemType = shp.Parent.Name & "(" & mSheetControls.CntrlType(shp) & ")": ServicedItemName = mSheetControls.CntrlName(shp)
        Case "Worksheet":       Set ws = v:     ServicedItemType = "Worksheet Name(CodeName)":                                  ServicedItemName = FinalName(ws.Name) & "(" & FinalName(ws.CodeName) & ")"
        Case "String"
            sServicedItemType = vbNullString
            sServicedItemName = vbNullString
            sServicedItem = mBasic.Align(v, lMaxLenServicedItemType + lMaxLenServicedItemName + 1, , " ", ".") & ": "
        Case "CommandButton":   Set cbt = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.Name
        Case "Image":           Set img = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.Name
        Case "Label":           Set lbl = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.Name
        Case "ListBox":         Set lbx = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.Name
        Case "OptionButton":    Set obt = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.Name
        Case "ScrollBar":       Set scb = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.Name
        Case "SpinButton":      Set sbt = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.Name
        Case "TextBox":         Set tbx = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.Name
        Case "ToggleButton":    Set tbt = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.Name
        Case "Range":           Set rng = v:    ServicedItemType = "Named Range":       ServicedItemName = rng.Name.Name
        Case Else
            Debug.Print "TypeName '" & TypeName(v) & "' not yet considered"
    End Select
End Property

Private Property Get ServicedItemName() As String
    ServicedItemName = sServicedItemName & "(" & lMaxLenServicedItemName & ")"
End Property

Private Property Let ServicedItemName(ByVal s As String)
    sServicedItemName = s
    lMaxLenServicedItemName = Max(lMaxLenServicedItemName, Len(sServicedItemName))
End Property

Private Property Get ServicedItemType() As String
    ServicedItemType = sServicedItemType & "(" & lMaxLenServicedItemType & ")"
End Property

Private Property Let ServicedItemType(ByVal s As String)
    sServicedItemType = s
    lMaxLenServicedItemType = Max(lMaxLenServicedItemType, Len(sServicedItemType))
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsLog" & "." & sProc
End Function

