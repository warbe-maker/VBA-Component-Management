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

Private Const SERVICES_LOG_FILE  As String = "CompMan.Services.log"

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

Friend Property Get Service(Optional ByVal new_log As Boolean = False) As String
    new_log = new_log ' not used
    Service = sService
End Property

Friend Property Let Service(Optional ByVal new_log As Boolean = False, _
                                     ByVal sv_name As String)
' ----------------------------------------------------------------------------
' - Assembles the service's Application.StatusBar message
' - Creates a new log file when none yet exists or when a new log is requested
'   (new_log = True)
' - Add a services split line when there are already logs in the file
' - Wirite a log header for the service
' ----------------------------------------------------------------------------
    Const PROC = "Service-Let"
    
    On Error GoTo eh
    
    sService = sv_name & " (by "
    If ThisWorkbook.name = mMe.CompManAddinName _
    Then sService = sService & "Addin" _
    Else sService = sService & ThisWorkbook.name
    sService = sService & ") for " & mService.Serviced.name & ": "
    
    bNewLog = new_log
    Me.Folder = mService.Serviced.Path
    
    '~~ In case no specific log-file had been decided a default log-file is determined
    If Me.File = vbNullString Then sFile = Me.Folder & "\" & SERVICES_LOG_FILE

    With New FileSystemObject
        If bNewLog And .FileExists(sFile) Then .DeleteFile sFile
        If Not .FileExists(sFile) Then .CreateTextFile (sFile)
        bNewLog = False
    End With
    
    If mFile.Txt(ft_file:=sFile) <> vbNullString Then
        '~~ Insert a new log section line when hte file already has written logs
        mFile.Txt(ft_file:=sFile _
                 , ft_append:=True _
                  ) = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.String$(Len(sService), "=")
    End If
    
    '~~ Write a service header
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
    Dim nm  As name
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
                Case vbext_ct_ActiveXDesigner:  ServicedItemType = "ActiveX-Designer":      ServicedItemName = vbc.name
                Case vbext_ct_ClassModule:      ServicedItemType = "Class-Module":          ServicedItemName = vbc.name
                Case vbext_ct_MSForm:           ServicedItemType = "UserForm":              ServicedItemName = vbc.name
                Case vbext_ct_StdModule:        ServicedItemType = "Standard-Module":       ServicedItemName = vbc.name
                Case vbext_ct_Document
                    If mComp.IsSheetDocMod(vbc) Then
                                                ServicedItemType = "Worksheet CodeName":    ServicedItemName = vbc.name
                    Else
                                                ServicedItemType = "Workbook":              ServicedItemName = vbc.name
                    End If
            End Select
        Case "Name"
            Set nm = v
            ServicedItemType = "Name"
            ServicedItemName = Replace(nm.name, nm.Parent.name & "!", vbNullString) & "(" & nm.RefersTo & ")"
            If MaxLenServicedItemName = 60 Then Stop

        Case "OLEObject":       Set oob = v:    ServicedItemType = "(Sheet)ActiveX":    ServicedItemName = "(" & oob.Parent.name & ")" & oob.name
        Case "Reference":       Set ref = v:    ServicedItemType = TypeName(ref):                                               ServicedItemName = ref.name
        Case "Shape":           Set shp = v:    ServicedItemType = shp.Parent.name & "(" & mSheetControls.CntrlType(shp) & ")": ServicedItemName = mSheetControls.CntrlName(shp)
        Case "Worksheet":       Set ws = v:     ServicedItemType = "Worksheet Name(CodeName)":                                  ServicedItemName = FinalName(ws.name) & "(" & FinalName(ws.CodeName) & ")"
        Case "String"
            sServicedItemType = vbNullString
            sServicedItemName = vbNullString
            sServicedItem = mBasic.Align(v, lMaxLenServicedItemType + lMaxLenServicedItemName + 1, , " ", ".") & ": "
        Case "CommandButton":   Set cbt = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.name
        Case "Image":           Set img = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.name
        Case "Label":           Set lbl = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.name
        Case "ListBox":         Set lbx = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.name
        Case "OptionButton":    Set obt = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.name
        Case "ScrollBar":       Set scb = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.name
        Case "SpinButton":      Set sbt = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.name
        Case "TextBox":         Set tbx = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.name
        Case "ToggleButton":    Set tbt = v:    ServicedItemType = TypeName(v):         ServicedItemName = v.name
        Case "Range":           Set rng = v:    ServicedItemType = "Named Range":       ServicedItemName = rng.name.name
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

