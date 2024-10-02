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
' LastModAtDateTimeUTC r/w
' LastModExpFile        r
' ----------------------------------------------------------------------------------------
Private fSavedExpFile           As File
Private sCompName               As String
Private sLastModByWbkFullName   As String
Private sLastModByWbkName       As String
Private sLastModAtDateTimeUTC      As String

Friend Property Let CompName(ByVal s As String)
    sCompName = s
End Property

Private Property Get LastModByWbkFullName() As String:          LastModByWbkFullName = sLastModByWbkFullName:                   End Property

Friend Property Get LastModByWbkName() As String:               LastModByWbkName = sLastModByWbkName:                           End Property

Friend Property Let LastModByWbkName(ByVal s As String):        sLastModByWbkName = s:                                          End Property

Friend Property Get LastModExpFile() As File:                   Set LastModExpFile = fSavedExpFile:                             End Property

Friend Property Let LastModExpFile(ByVal f As File):            Set fSavedExpFile = f:                                          End Property

Friend Property Get LastModAtDateTimeUTC() As String:           LastModAtDateTimeUTC = sLastModAtDateTimeUTC:                   End Property

Friend Property Let LastModAtDateTimeUTC(ByVal s As String):    sLastModAtDateTimeUTC = s:                                      End Property

Friend Property Get LastModWbkBaseName(Optional ByVal l_comp As String = vbNullString) As String
    Const PROC = "LastModWbkBaseName"
    
    If l_comp = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No component name provided!"
    
    LastModWbkBaseName = CommonPublic.LastModInWbkBaseName(l_comp)
    
End Property

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "clsRaw." & s
End Function

