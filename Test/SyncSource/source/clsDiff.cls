VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lDiffLine   As Long
Private sLineLeft   As String
Private sLineRight  As String
Private sFileLeft   As String
Private sFileRight  As String

'Public Property Get DiffLine() As Long:             DiffLine = lDiffLine:   End Property
Public Property Let DiffLine(ByVal l As Long):      lDiffLine = l:          End Property

'Public Property Get LineLeft() As String:           LineLeft = sLineLeft:   End Property
Public Property Let LineLeft(ByVal s As String):    sLineLeft = s:          End Property

'Public Property Get LineRight() As String:          LineRight = sLineRight: End Property
Public Property Let LineRight(ByVal s As String):   sLineRight = s:         End Property

'Public Property Get FileLeft() As String:           FileLeft = sFileLeft:   End Property
Public Property Let FileLeft(ByVal s As String):    sFileLeft = s:          End Property

'Public Property Get FileRight() As String:          FileRight = sFileRight: End Property
Public Property Let FileRight(ByVal s As String):   sFileRight = s:         End Property

Public Function DiffItem() As String
' --------------------------------------------------------------------
'
' --------------------------------------------------------------------
    Dim i   As Long
    Dim sF1 As String
    Dim sF2 As String
    
    For i = 1 To mBasic.Min(Len(sFileLeft), Len(sFileRight))
        If VBA.Mid$(sFileLeft, i, 1) <> VBA.Mid$(sFileRight, i, 1) _
        Then Exit For
    Next i
    i = i - 2
    sF1 = "..." & VBA.Right$(sFileLeft, Len(sFileLeft) - i) & "Line " & Format(lDiffLine, "0000") & ": "
    sF2 = "..." & VBA.Right$(sFileRight, Len(sFileRight) - i) & "Line " & Format(lDiffLine, "0000") & ": "
    
    DiffItem = sF1 & "'" & sLineLeft & "'" & vbLf & sF2 & "'" & sFileRight & "'"

End Function


