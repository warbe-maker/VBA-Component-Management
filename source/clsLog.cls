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
' -----------------------------------------------------------------------------------
' Class Module clsLog
'
' Public services: - Entry                  Appends an entry to the log file.
'                  - FileFullName
'                  - Folder
'                  - LogFile
'                  - ServicedItemsMaxLen
'                  - Service
'                  - ServicedItem
' -----------------------------------------------------------------------------------
Private bNewLog                     As Boolean
Private sEntry                      As String
Private sFileFullName               As String
Private sService                    As String
Private sServicedItem               As String
Private sServicedItemName           As String
Private sServicedItemType           As String

Private Sub Class_Initialize()
    bNewLog = True
End Sub

Friend Property Let Entry(ByVal s As String)
' -----------------------------------------------------------------------------------
' Append an entry to the log file.
' -----------------------------------------------------------------------------------
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
    mFso.FileTxt(ft_file:=sFileFullName, ft_append:=True) = sEntry
    sItem = sServicedItem
    sCompName = Split(ServicedItemName, "(")(0)
    
End Property

Friend Property Get FileFullName() As String
' -----------------------------------------------------------------------------------
' -----------------------------------------------------------------------------------
    If sFileFullName = vbNullString Then
        sFileFullName = wsService.CurrentServiceLogFileFullName
    End If
    FileFullName = sFileFullName
End Property

Friend Property Let FileFullName(ByVal s As String)
    Dim fso As New FileSystemObject
    With fso
        If Not .FileExists(s) Then .CreateTextFile s
    End With
    sFileFullName = s
End Property

Friend Property Get LogFile() As File
' -----------------------------------------------------------------------------------
' Returns the log file as file object
' -----------------------------------------------------------------------------------
    With New FileSystemObject
        If Not .FileExists(sFileFullName) Then .CreateTextFile sFileFullName
        Set LogFile = .GetFile(sFileFullName)
    End With

End Property

Friend Property Let Service(Optional ByVal new_log As Boolean = False, _
                                     ByVal sv_name As String)
' ----------------------------------------------------------------------------
' - Assembles the service's Application.StatusBar message
' - Creates a new log file when none yet exists or when a new log is requested
'   (new_log = True) and the existing's creation dates back more than one day
' - Appends a service header (split line) when there are already logs in the
'   file and the service name has changed.
' - Wirite a log header for the service
' ----------------------------------------------------------------------------
    Const PROC = "Service-Let"
    
    On Error GoTo eh
    Dim bNewService As Boolean
    Dim sNewService As String
    Dim fso         As New FileSystemObject
    
    sNewService = sv_name & " (by "
    If ThisWorkbook.Name = mAddin.WbkName _
    Then sNewService = sNewService & "Add-in" _
    Else sNewService = sNewService & ThisWorkbook.Name
    sNewService = sNewService & ") for " & mService.Serviced.Name & ": "
    bNewService = sNewService <> sService
    sService = sNewService
    
    bNewLog = new_log
    '~~ In case no specific log-file had been decided a default log-file is determined
    If sFileFullName = vbNullString Then sFileFullName = mService.Serviced.Path & "\" & LOG_FILE_NAME

    With fso
        If bNewLog And .FileExists(sFileFullName) Then
            If VBA.DateDiff("d", .GetFile(sFileFullName).DateCreated, Now()) > 0 Then
                .DeleteFile sFileFullName
            End If
        End If
        If Not .FileExists(sFileFullName) Then .CreateTextFile sFileFullName
        bNewLog = False
    End With
    
    If bNewService Then
        '~~ When the service differs from a previous service
        If mFso.FileTxt(ft_file:=sFileFullName) <> vbNullString Then
            '~~ When the log file already contains log entries insert a split line
            mFso.FileTxt(ft_file:=sFileFullName _
                     , ft_append:=True _
                      ) = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.String$(Len(sService), "=")
        End If
    
        '~~ Append a service header to the log file
        mFso.FileTxt(ft_file:=sFileFullName _
                , ft_append:=True _
                 ) = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & sService & vbCrLf & _
                     Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.String$(Len(sService), "-")
    End If

xt: wsService.CurrentServiceLogFileFullName = sFileFullName
    Set fso = Nothing
    Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get ServicedItem() As Variant
    If sServicedItemType = vbNullString And sServicedItemName = vbNullString Then
        ServicedItem = sServicedItem
    Else
        sServicedItem = mBasic.Align(sServicedItemType, wsService.ServicedItemsMaxLenType) & " " & mBasic.Align(sServicedItemName, wsService.ServicedItemsMaxLenName, , " ", ".") & ": "
        ServicedItem = sServicedItem
    End If
End Property

Friend Property Let ServicedItem(ByVal v As Variant)
' -----------------------------------------------------------------------------------
' Determines the serviced item's type and name string for logging and computes the
' maximum length for both.
' -----------------------------------------------------------------------------------
    Const PROC = "ServicedItem-Let"
    
    On Error GoTo eh
    Dim vbc As VBComponent
    Dim nme  As Name
    Dim wsh  As Worksheet
    Dim ref As Reference
    Dim shp As Shape
    Dim tbt As ToggleButton
    Dim tbx As TextBox
    Dim sbt As SpinButton
    Dim scb As ScrollBar
    Dim obt As OptionButton
    Dim lbx As ListBox
    Dim lbl As Label
    Dim img As Image
    Dim cbt As CommandButton
    Dim Rng As Range
    
    Select Case TypeName(v)
        Case "VBComponent":     Set vbc = v:    Select Case vbc.Type
                                                    Case vbext_ct_ActiveXDesigner:  ServicedItemProperties "ActiveX-Designer", vbc.Name
                                                    Case vbext_ct_ClassModule:      ServicedItemProperties "Class-Module", vbc.Name
                                                    Case vbext_ct_MSForm:           ServicedItemProperties "UserForm", vbc.Name
                                                    Case vbext_ct_StdModule:        ServicedItemProperties "Standard-Module", vbc.Name
                                                    Case vbext_ct_Document:         If mComp.IsSheetDocMod(vbc, mService.Serviced) _
                                                                                    Then ServicedItemProperties "Worksheet", vbc.Name _
                                                                                    Else ServicedItemProperties "Workbook", vbc.Name
                                                End Select
        Case "Name":            Set nme = v:    ServicedItemProperties "Name", Replace(nme.Name, nme.Parent.Name & "!", vbNullString) & "(" & nme.RefersTo & ")"
        Case "Reference":       Set ref = v:    ServicedItemProperties TypeName(ref), ref.Description
        Case "Shape":           Set shp = v:    ServicedItemProperties mSyncShapes.TypeString(shp), shp.Parent.Name & "." & mSyncShapes.ShapeNames(shp)
        Case "Worksheet":       Set wsh = v:    ServicedItemProperties "Worksheet", wsh.CodeName
        Case "String":                          ServicedItemProperties Split(v, ";")(0), Split(v, ";")(1)
        Case "CommandButton":   Set cbt = v:    ServicedItemProperties TypeName(v), v.Name
        Case "Image":           Set img = v:    ServicedItemProperties TypeName(v), v.Name
        Case "Label":           Set lbl = v:    ServicedItemProperties TypeName(v), v.Name
        Case "ListBox":         Set lbx = v:    ServicedItemProperties TypeName(v), v.Name
        Case "OptionButton":    Set obt = v:    ServicedItemProperties TypeName(v), v.Name
        Case "ScrollBar":       Set scb = v:    ServicedItemProperties TypeName(v), v.Name
        Case "SpinButton":      Set sbt = v:    ServicedItemProperties TypeName(v), v.Name
        Case "TextBox":         Set tbx = v:    ServicedItemProperties TypeName(v), v.Name
        Case "ToggleButton":    Set tbt = v:    ServicedItemProperties TypeName(v), v.Name
        Case "Range":           Set Rng = v:    ServicedItemProperties TypeName(v), v.Name.Name
        Case Else
            Debug.Print "TypeName '" & TypeName(v) & "' not yet considered as a serviced item"
    End Select

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get ServicedItemName() As String
    ServicedItemName = sServicedItemName & "(" & wsService.ServicedItemsMaxLenName & ")"
End Property

Private Property Let ServicedItemName(ByVal s As String)
    sServicedItemName = s
End Property

Private Property Get ServicedItemType() As String
    ServicedItemType = sServicedItemType & "(" & wsService.ServicedItemsMaxLenType & ")"
End Property

Private Property Let ServicedItemType(ByVal s As String)
    sServicedItemType = s
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsLog" & "." & sProc
End Function

Private Sub ServicedItemProperties(ByVal sip_type As String, _
                                   ByVal sip_name As String)
    ServicedItemType = sip_type
    ServicedItemName = sip_name
    wsService.ServicedItemsMaxLenType = Max(wsService.ServicedItemsMaxLenType, Len(sip_type))
    wsService.ServicedItemsMaxLenName = Max(wsService.ServicedItemsMaxLenName, Len(sip_name))

End Sub

