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
Private sTempFolder                 As String
Private wbRawHost                   As Workbook
Private sRawHostWbFullName          As String
Private sHostWbPath                 As String
Private sExpFileExt                 As String
Private vbcRaw                      As VBComponent

Private Sub Class_Terminate()
    CleanUpTemps
    Set fso = Nothing
End Sub

Public Property Let CompName(ByVal s As String):        sCompName = s:                                                              End Property

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

Public Property Get ExpFileTemp() As File
' ------------------------------------------------------------------------------
' Returns a temporary Export File which represents an open Raw's current code -
' which may yet not have been exported because the host has yet not been saved.
' ------------------------------------------------------------------------------
    Set ExpFileTemp = fso.GetFile(Me.ExpFileTempFullName)
End Property

Public Property Get ExpFileTempFullName() As String
' ------------------------------------------------------------------------------
' Returns the full name for a temporary Export File for an open Raw Host
' Workbook. This property is only available for an open Raw Host. It is provided
' along with the HostIsOpen property.
' ------------------------------------------------------------------------------
    Const PROC = "ExpFileTempFullName-Get"
    
    On Error GoTo eh
    With fso
        sTempFolder = mExport.ExpFileFolderPath(Me.RawHostWbFullName) & "\Temp"
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
    
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get HostIsOpen() As Boolean
' ------------------------------------------------------------------------------
' Note: This property must have been requested and returned TRUE when the
' following properties provided herewith are obtained:
' - HostwbPath (sHostWbPath)
' - VBComp     (vbcRaw
' - RawHost    (wbRawHost)
' ------------------------------------------------------------------------------
    HostIsOpen = mWrkbk.IsOpen(Me.RawHostWbName)
    If HostIsOpen Then
        Set wbRawHost = mWrkbk.GetOpen(Me.RawHostWbName)
        Set vbcRaw = wbRawHost.VBProject.VBComponents(sCompName)
        sHostWbPath = wbRawHost.Path
    End If
End Property

Public Property Get HostWbPath() As String:         HostWbPath = sHostWbPath:                                               End Property

Public Property Get RawHostWbBaseName() As String:      RawHostWbBaseName = mComCompsRawsSaved.RawHostWbBaseName(sCompName):        End Property

Public Property Get RawHostWbFullName() As String
' ------------------------------------------------------------------------------
' Returns the full name of the Raw's host which has been saved along with the
' last open of the Workbook. When this Workbook is open (property HostIsOpen)
' this is identical with its FullName
' ------------------------------------------------------------------------------
    If sRawHostWbFullName = vbNullString Then
        sRawHostWbFullName = mComCompsRawsSaved.RawHostWbFullName(sCompName)
    End If
    RawHostWbFullName = sRawHostWbFullName
End Property

Public Property Get RawHostWbName() As String:          RawHostWbName = mComCompsRawsSaved.RawHostWbName(sCompName):                End Property

Public Property Get RevisionNumber() As String:     RevisionNumber = mComCompsRawsSaved.RawSavedRevisionNumber(sCompName):  End Property

Public Property Get SavedExpFile() As File
' ------------------------------------------------------------------------------
' Returns the Raws's Export File which had been saved to the Common Components
' folder
' ------------------------------------------------------------------------------
    Set SavedExpFile = fso.GetFile(mComCompsRawsSaved.SavedExpFileFullName(sCompName))
End Property

Public Property Get SavedExpFileFullName() As String:   SavedExpFileFullName = mComCompsRawsSaved.SavedExpFileFullName(sCompName):  End Property

Public Property Get VBComp() As VBComponent:        Set VBComp = vbcRaw:                                                    End Property

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
