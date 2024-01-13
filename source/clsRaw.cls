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
' LastModAtDateTime r/w
' LastModExpFile        r
' ----------------------------------------------------------------------------------------
Private fSavedExpFile           As File
Private sCompName               As String
Private sLastModByWbkFullName   As String
Private sLastModByWbkName       As String
Private sLastModAtDateTime      As String

Friend Property Let CompName(ByVal s As String)
    sCompName = s
    CommComps.CompName = s
End Property

Private Property Get LastModByWbkFullName() As String:          LastModByWbkFullName = sLastModByWbkFullName:                   End Property

Friend Property Get LastModByWbkName() As String:               LastModByWbkName = sLastModByWbkName:                           End Property

Friend Property Let LastModByWbkName(ByVal s As String):        sLastModByWbkName = s:                                          End Property

Friend Property Get LastModExpFile() As File:                   Set LastModExpFile = fSavedExpFile:                             End Property

Friend Property Let LastModExpFile(ByVal f As File):            Set fSavedExpFile = f:                                          End Property

Friend Property Get LastModAtDateTime() As String:               LastModAtDateTime = sLastModAtDateTime:                           End Property

Friend Property Let LastModAtDateTime(ByVal s As String):        sLastModAtDateTime = s:                                          End Property

Friend Property Get LastModWbkBaseName() As String:             LastModWbkBaseName = CommComps.LastModInWbkBaseName:            End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsRaw" & "." & sProc
End Function
