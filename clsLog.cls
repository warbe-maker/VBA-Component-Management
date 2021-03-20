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

Private wbServiced              As Workbook
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

Friend Property Get MaxLenServicedItemType() As Long
    MaxLenServicedItemType = lMaxLenServicedItemType
End Property

Friend Property Get MaxLenServicedItemName() As Long
    MaxLenServicedItemName = lMaxLenServicedItemName
End Property

Friend Property Let Entry(ByVal s As String)
' ------------------------------------------
' Append an entry to the log file.
' ------------------------------------------
Static sItem As String
  
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
    
End Property

Friend Property Get File() As String:               File = sFile:               End Property

Friend Property Let File(ByVal s As String):        sFile = s:                  End Property

Friend Property Get Folder() As String:             Folder = sFolder:           End Property

Friend Property Let Folder(ByVal s As String):      sFolder = s:                End Property

'Friend Property Let Info(ByVal s As String)
'' ------------------------------------------
'' Append an info line to the log file.
'' ------------------------------------------
'Static sItem As String
'
'    If sServicedItem = vbNullString Then
'        sEntry = s
'    Else
'        If sServicedItem = sItem Then
'            sEntry = VBA.Space$(Len(sServicedItem) + 2) & s
'        Else
'            sEntry = sServicedItem & ": " & s
'        End If
'    End If
'    mFile.Txt(ft_file:=sFile, ft_append:=True) = sEntry
'    sItem = sServicedItem
'
'End Property

Friend Property Get LogFile() As File
    Dim sLogFullName    As String
    sLogFullName = Me.Folder & "\" & Me.File
    With New FileSystemObject
        If Not .FileExists(sLogFullName) Then .CreateTextFile sLogFullName
        Set LogFile = .GetFile(sLogFullName)
    End With
End Property

Friend Property Get Service() As String
    Service = sService
End Property

Friend Property Let Service(ByVal sv_name As String)

    sService = sv_name & " by " & ThisWorkbook.Name & " for '" & wbServiced.Name & "': "
    
    If mFile.Txt(ft_file:=sFile) <> vbNullString _
    Then mFile.Txt(ft_file:=sFile _
                 , ft_append:=True _
                  ) = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.String$(Len(sService), "=")
    
    mFile.Txt(ft_file:=sFile _
            , ft_append:=True _
             ) = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & sService & vbCrLf & _
                 Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.String$(Len(sService), "-")
End Property

Friend Property Get ServicedItem() As Variant
    If sServicedItemType = vbNullString And sServicedItemName = vbNullString Then
        ServicedItem = sServicedItem
    Else
        sServicedItem = Align(sServicedItemType, lMaxLenServicedItemType) & " " & Align(sServicedItemName, lMaxLenServicedItemName, , " ", ".") & ": "
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
    
    Select Case TypeName(v)
        Case "VBComponent"
            Set vbc = v
            Select Case vbc.Type
                Case vbext_ct_ActiveXDesigner:  ServicedItemType = "ActiveX-Designer":  ServicedItemName = vbc.Name
                Case vbext_ct_ClassModule:      ServicedItemType = "Class-Module":      ServicedItemName = vbc.Name
                Case vbext_ct_MSForm:           ServicedItemType = "UserForm":          ServicedItemName = vbc.Name
                Case vbext_ct_StdModule:        ServicedItemType = "Standard-Module":   ServicedItemName = vbc.Name
                Case vbext_ct_Document
                    If mCompMan.IsSheetComp(vbc) Then
                                                ServicedItemType = "Worksheet":         ServicedItemName = vbc.Name
                    Else
                                                ServicedItemType = "Workbook":          ServicedItemName = vbc.Name
                    End If
            End Select
        Case "Name":        Set nm = v:         ServicedItemType = "Name":              ServicedItemName = nm.Name & "(" & nm.RefersTo & ")"
        Case "Reference":   Set ref = v:        ServicedItemType = TypeName(ref):       ServicedItemName = ref.Name
        Case "Shape":       Set shp = v:        ServicedItemType = "Shape":             ServicedItemName = shp.Name
        Case "Worksheet":   Set ws = v:         ServicedItemType = "Worksheet":         ServicedItemName = ws.Name & "(" & ws.CodeName & ")"
        Case "String"
            sServicedItemType = vbNullString
            sServicedItemName = vbNullString
            sServicedItem = Align(v, lMaxLenServicedItemType + lMaxLenServicedItemName + 1, , " ", ".") & ": "
    End Select
       
End Property

Private Property Let ServicedItemType(ByVal s As String)
    Dim lMax As Long
    sServicedItemType = s
    lMaxLenServicedItemType = Max(lMaxLenServicedItemType, Len(sServicedItemType))
End Property

Private Property Let ServicedItemName(ByVal s As String)
    Dim lMax As Long
    sServicedItemName = s
    lMaxLenServicedItemName = Max(lMaxLenServicedItemName, Len(sServicedItemName))
End Property

Friend Property Set ServicedWrkbk( _
                   Optional ByVal sw_new_log As Boolean = False, _
                            ByRef sw_wbServiced As Workbook)
' ----------------------------------------------------------------
'
' ----------------------------------------------------------------
    Dim vbc As VBComponent
    
    bNewLog = sw_new_log
    Set wbServiced = sw_wbServiced
    Me.Folder = sw_wbServiced.Path
    
    If sFile = vbNullString Then
        '~~ In case not a specific file had explicitely assigned before
        sFile = Me.Folder & "\" & mService.SERVICES_LOG_FILE
    End If
    With New FileSystemObject
        If bNewLog Then
            If .FileExists(sFile) Then .DeleteFile sFile
            .CreateTextFile (sFile)
            bNewLog = False
        End If
    End With

End Property

