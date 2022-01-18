VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRaw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------------------------
' Class Module clsRaw
' Provides all Properties for the Raw of a Used Common Component whereby all Properties
' are based on the CompName property assigned when the class object is established.
'
' Properties:
' CompName
'
' ----------------------------------------------------------------------------------------
Private fso                         As New FileSystemObject
Private sCompName                   As String
Private sExpFileTempFullName        As String
Private sMostRecentExpFileFullName  As String
Private sTempFolder                 As String
Private wbRawHost                   As Workbook
Private sHostWbFullName             As String
Private sHostWbPath                 As String
Private sExpFileExt                 As String
Private vbcRaw                      As VBComponent

Private Sub Class_Terminate()
    CleanUpTemps
    Set fso = Nothing
End Sub

Public Property Let CompName(ByVal s As String):    sCompName = s:                                                  End Property

Public Property Get ExpFileFullName() As String:    ExpFileFullName = mComCompsRawsSaved.ExpFileFullName(sCompName):    End Property

Public Property Get HostWbBaseName() As String:     HostWbBaseName = mComCompsRawsSaved.HostWbBaseName(sCompName):   End Property

Public Property Get HostWbFullName() As String
' ------------------------------------------------------------------------------
' Returns the full name of the Raw's host which has been saved along with the
' last open of the Workbook. When this Workbook is open (property HostIsOpen)
' this is identical with its FullName
' ------------------------------------------------------------------------------
    sHostWbFullName = mComCompsRawsSaved.HostWbFullName(sCompName)
    HostWbFullName = sHostWbFullName
End Property

Public Property Get HostWbName() As String:         HostWbName = mComCompsRawsSaved.HostWbName(sCompName):           End Property

Public Property Get ExpFile() As File
' ------------------------------------------------------------------------------
' Returns the Raws's Export File which had been saved to the Common Components
' folder
' ------------------------------------------------------------------------------
    Set ExpFile = fso.GetFile(mComCompsRawsSaved.ExpFileFullName(sCompName))
End Property

Public Property Get ExpFileTemp() As File
' ------------------------------------------------------------------------------
' Returns a temporary Export File which represents an open Raw's current code -
' which may yet not have been exported because the host has yet not been saved.
' ------------------------------------------------------------------------------
    Set ExpFileTemp = fso.GetFile(Me.ExpFileTempFullName)
End Property

Friend Property Get ExpFileExt() As String
    
    If sExpFileExt = vbNullString Then
        If vbcRaw Is Nothing Then Set vbcRaw = wbRawHost.VBProject.VBComponents(sCompName)
        With Me
            Select Case vbcRaw.Type
                Case vbext_ct_StdModule:    sExpFileExt = ".bas"
                Case vbext_ct_ClassModule:  sExpFileExt = ".cls"
                Case vbext_ct_MSForm:       sExpFileExt = ".frm"
                Case vbext_ct_Document:     sExpFileExt = ".cls"
            End Select
        End With
    End If
    ExpFileExt = sExpFileExt

End Property

Public Property Get ExpFileTempFullName() As String
' ------------------------------------------------------------------------------
' Returns the full name for a temporary Export File for an open Raw Host
' Workbook. This property is only available for an open Raw Host. It is provided
' along with the HostIsOpen property.
' ------------------------------------------------------------------------------
    
    With fso
        sTempFolder = mExport.ExpFileFolderPath(sHostWbFullName) & "\Temp"
        If Not .FolderExists(sTempFolder) Then
            DoEvents
            .CreateFolder sTempFolder
            DoEvents
        End If
        ExpFileTempFullName = sTempFolder & "\" & sCompName & Me.ExpFileExt
        If Not .FileExists(ExpFileTempFullName) Then
            vbcRaw.Export ExpFileTempFullName
        End If
    End With
    
End Property

Public Property Get HostIsOpen() As Boolean
' ------------------------------------------------------------------------------
' Note: This property must have been requested and returned TRUE when the
' following properties provided herewith are obtained:
' - HostwbPath (sHostWbPath)
' - VBComp     (vbcRaw
' - RawHost    (wbRawHost)
' ------------------------------------------------------------------------------
    HostIsOpen = mWrkbk.IsOpen(Me.HostWbName)
    If HostIsOpen Then
        Set wbRawHost = mWrkbk.GetOpen(Me.HostWbName)
        Set vbcRaw = wbRawHost.VBProject.VBComponents(sCompName)
        sHostWbPath = wbRawHost.Path
    End If
End Property
Public Property Get HostWbPath() As String:         HostWbPath = sHostWbPath:                                       End Property
Public Property Get VBComp() As VBComponent:        Set VBComp = vbcRaw:                                            End Property
Public Property Get RevisionNumber() As String:     RevisionNumber = mComCompsRawsSaved.RevisionNumber(sCompName):   End Property

Private Sub CleanUpTemps()
    With fso
        If .FolderExists(sTempFolder) Then
            .GetFolder(sTempFolder).Delete
        End If
    End With
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsRaw" & "." & sProc
End Function

'Friend Property Get Changed( _
'                      ByRef clone As clsComp, _
'             Optional ByVal check_again As Boolean = False) As Boolean
'' --------------------------------------------------------------------
'' Returns TRUE when the raw's regular Export-File differs from the
'' clone's regular Export-File which indicates that the code of the raw
'' had changed.
'' Expception: When the clone's coresponding Export-File does not exist
''             a change is regarded in the sense that the raw's Export-
''             File indicates a new component
'' --------------------------------------------------------------------
'    Const PROC = "Changed-Get"
'
'    On Error GoTo eh
'    Dim fso As New FileSystemObject
'
'    If dctChanges Is Nothing Or check_again Then
'        '~~ The difference between the raw's and the clone's
'        '~~ export file provided only oncce or when explicitely requested
'        If Not fso.FileExists(clone.ExpFileFullName) Then
'            '~~ The component never had been exported yet or the Export-File has been deleted
'            clone.Export
'        End If
'        Set dctChanges = _
'        mService.FilesDifference(fd_exp_file_1:=fso.GetFile(clone.ExpFileFullName) _
'                               , fd_exp_file_2:=fso.GetFile(Me.ExpFileFullName))
'    End If
'    Changed = dctChanges.Count <> 0
'
'xt: Set fso = Nothing
'    Exit Property
'
'eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Property


