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
' ---------------------------------------------------------
' Class Module clsRaw Properties of a remote raw component.
'
' W. Rauschenberger Berlin, Jan 2021 (last revised)
' ---------------------------------------------------------
Private wbRawHost           As Workbook
Private sHostFullName       As String
Private sExpFileFullName    As String
Private flExpFile           As FILE
Private sCompName           As String
Private vbcRaw              As VBComponent

Public Property Get CodeAsOfDate() As String:           CodeAsOfDate = flExpFile.DateLastModified:      End Property

Public Property Get CompName() As String:               CompName = sCompName:                           End Property

Public Property Let CompName(ByVal s As String):        sCompName = s:                                  End Property

Public Property Get ExpFile() As FILE:                  Set ExpFile = flExpFile:                        End Property

Public Property Let ExpFile(ByVal fl As FILE):          Set flExpFile = fl:                             End Property

Public Property Get ExpFileFullName() As String:        ExpFileFullName = sExpFileFullName:             End Property

Public Property Let ExpFileFullName(ByVal s As String): sExpFileFullName = s:                           End Property

Public Property Get HostFullName() As String:           HostFullName = sHostFullName:                   End Property

Public Property Let HostFullName(ByVal s As String):    sHostFullName = s:                              End Property

Public Property Get RawHost() As Workbook:              Set RawHost = wbRawHost:                        End Property

Public Property Let RawHost(ByVal wb As Workbook):      Set wbRawHost = wb:                             End Property

Public Property Get RawType() As vbext_ComponentType:   RawType = vbcRaw.Type:                          End Property

Public Property Get RawVbc() As VBComponent:            Set RawVbc = vbcRaw:                            End Property

Public Property Let RawVbc(ByVal vbc As VBComponent):   Set vbcRaw = vbc:                               End Property

Public Property Get WrkbkBaseName() As String
    Dim fso As New FileSystemObject
    WrkbkBaseName = fso.GetBaseName(sHostFullName)
    Set fso = Nothing
End Property
