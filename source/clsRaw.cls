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
Private fso                 As New FileSystemObject
Private sCompName           As String
Private sTempFolder         As String
Private sRawHostWbFullName  As String
Private sExpFileExt         As String
Private vbcRaw              As VBComponent

Private Sub Class_Terminate()
    CleanUpTemps
    Set fso = Nothing
End Sub

Public Property Let CompName(ByVal s As String):        sCompName = s:                                                              End Property

Private Property Let ExpFileExt(ByVal s As String):      sExpFileExt = s:                                                    End Property

Private Property Get ExpFileExt() As String:             ExpFileExt = sExpFileExt:                                           End Property

Public Property Get ExpFileTemp() As File
' ------------------------------------------------------------------------------
' Returns a temporary Export File which represents an open Raw's current code -
' which may yet not have been exported because the host has yet not been saved.
' ------------------------------------------------------------------------------
    Set ExpFileTemp = fso.GetFile(ExpFileTempFullName)
End Property

Private Property Get ExpFileTempFullName() As String
' ------------------------------------------------------------------------------
' Returns the full name for a temporary Export File for an open Raw Host
' Workbook. This property is only available for an open Raw Host. It is provided
' along with the HostIsOpen property.
' ------------------------------------------------------------------------------
    Const PROC = "ExpFileTempFullName-Get"
    
    On Error GoTo eh
    With fso
        sTempFolder = mExport.ExpFileFolderPath(RawHostWbFullName) & "\Temp"
        If Not .FolderExists(sTempFolder) Then
            mBasic.TimedDoEvents ErrSrc(PROC)
            .CreateFolder sTempFolder
            mBasic.TimedDoEvents ErrSrc(PROC)
        End If
        ExpFileTempFullName = sTempFolder & "\" & sCompName & ExpFileExt
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

Private Property Get RawHostWbBaseName() As String:      RawHostWbBaseName = mCommComps.RawHostWbBaseName(sCompName):        End Property

Private Property Get RawHostWbFullName() As String
' ------------------------------------------------------------------------------
' Returns the full name of the Raw's host which has been saved along with the
' last open of the Workbook. When this Workbook is open (property HostIsOpen)
' this is identical with its FullName
' ------------------------------------------------------------------------------
    If sRawHostWbFullName = vbNullString Then
        sRawHostWbFullName = mCommComps.RawHostWbFullName(sCompName)
    End If
    RawHostWbFullName = sRawHostWbFullName
End Property

Private Property Get RawHostWbName() As String:         RawHostWbName = mCommComps.RawHostWbName(sCompName):                End Property

Friend Property Get RevisionNumber() As String:         RevisionNumber = mCommComps.RevisionNumber(sCompName):              End Property

Public Property Get SavedExpFile() As File
' ------------------------------------------------------------------------------
' Returns the Raws's Export File which had been saved to the Common Components
' folder
' ------------------------------------------------------------------------------
    Set SavedExpFile = fso.GetFile(mCommComps.SavedExpFileFullName(sCompName))
End Property

Public Property Get SavedExpFileFullName() As String:   SavedExpFileFullName = mCommComps.SavedExpFileFullName(sCompName):  End Property

Private Property Get VBComp() As VBComponent:           Set VBComp = vbcRaw:                                                End Property

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
