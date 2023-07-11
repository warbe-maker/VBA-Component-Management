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
' Class Module clsRaw: Provides all properties a Common Component in the Common-Components
' ==================== folder.
'
' Public Properties:
' ------------------
' CompName                w
' LastModByWbkName      r/w
' LastModWbkBaseName    r
' LastModRevisionNumber r/w
' LastModExpFile        r
' ----------------------------------------------------------------------------------------
Private fSavedExpFile           As File
Private fso                     As FileSystemObject
Private sCompName               As String
Private sLastModByWbkFullName   As String
Private sLastModByWbkName       As String
Private sLastModRevisionNumber  As String

Private Sub Class_Initialize()
    Set fso = New FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set fso = Nothing
End Sub

Friend Property Let CompName(ByVal s As String):                sCompName = s:                                                  End Property

Private Property Get LastModByWbkFullName() As String:          LastModByWbkFullName = sLastModByWbkFullName:                   End Property

Friend Property Get LastModByWbkName() As String:               LastModByWbkName = sLastModByWbkName:                           End Property

Friend Property Let LastModByWbkName(ByVal s As String):        sLastModByWbkName = s:                                          End Property

Friend Property Get LastModExpFile() As File:                   Set LastModExpFile = fSavedExpFile:                             End Property

Friend Property Let LastModExpFile(ByVal f As File):            Set fSavedExpFile = f:                                          End Property

Friend Property Get LastModRevisionNumber() As String:          LastModRevisionNumber = sLastModRevisionNumber:                 End Property

Friend Property Let LastModRevisionNumber(ByVal s As String):   sLastModRevisionNumber = s:                                     End Property

Friend Property Get LastModWbkBaseName() As String:             LastModWbkBaseName = CommComps.LastModWbkBaseName(sCompName):  End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsRaw" & "." & sProc
End Function
