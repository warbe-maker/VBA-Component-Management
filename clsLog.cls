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

Private wbServiced          As Workbook
Private sService            As String
Private sServicedItem       As String
Private sFile               As String
Private sFolder             As String
Private lCompMaxLen         As Long
Private bNewLog             As Boolean
Private sEntry              As String

Private Sub Class_Initialize()
    bNewLog = True
End Sub

Friend Property Get CompMaxLen() As Long:           CompMaxLen = lCompMaxLen:   End Property

Friend Property Let CompMaxLen(ByVal l As Long):    lCompMaxLen = l:            End Property

Friend Property Let Entry(ByVal s As String)
' ------------------------------------------
' Append an entry to the log file.
' ------------------------------------------
Static sItem As String

    If sServicedItem = vbNullString Then
        sEntry = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & s
    Else
        If sServicedItem = sItem Then
            sEntry = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.Space$(Len(sServicedItem) + 2) & s
        Else
            sEntry = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & sServicedItem & ": " & s
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

    sService = ThisWorkbook.Name & " '" & sv_name & "' for '" & wbServiced.Name & "'"
    
    If mFile.Txt(ft_file:=sFile) <> vbNullString _
    Then mFile.Txt(ft_file:=sFile _
                 , ft_append:=True _
                  ) = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.String$(Len(sService), "=")
    
    mFile.Txt(ft_file:=sFile _
            , ft_append:=True _
             ) = Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & sService & vbCrLf & _
                 Format$(Now(), "YY-MM-DD hh:mm:ss") & " " & VBA.String$(Len(sService), "-")
End Property

Friend Property Let ServicedItem(ByVal s As String)
    sServicedItem = s & " " & String(lCompMaxLen - Min(lCompMaxLen, Len(s)), ".")
End Property

Friend Property Set ServicedWrkbk( _
                   Optional ByVal sw_new_log As Boolean = False, _
                            ByRef sw_wbServiced As Workbook)
    
    Dim vbc As VBComponent
    
    bNewLog = sw_new_log
    Set wbServiced = sw_wbServiced
    Me.Folder = sw_wbServiced.Path
    If lCompMaxLen = 0 Then
        For Each vbc In wbServiced.VBProject.VBComponents
            lCompMaxLen = mBasic.Max(lCompMaxLen, Len(vbc.Name))
        Next vbc
    End If
    
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

