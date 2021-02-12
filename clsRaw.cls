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
' ------------------------------------------------------------
' Class Module clsRaw Properties of a raw component. I.e.
'       a component a certain Workbook claims the raw version.
'
' W. Rauschenberger Berlin, Jan 2021 (last revised)
' -----------------------------------------------------------
Private dctChanges          As Dictionary
Private flExpFile           As File
Private sCloneExpFileFullName   As String
Private sCompName           As String
Private sExpFileExtension   As String
Private sExpFileFullName    As String
Private sExpFilePath        As String
Private sHostFullName       As String
Private vbcRaw              As VBComponent
Private wbRawHost           As Workbook
Private sService            As String

Private Sub Class_Terminate()
    Set dctChanges = Nothing
End Sub

Friend Property Get Service() As String
    If sService = vbNullString Then
        sService = Split(Me.ExpFilePath, "\")(UBound(Split(Me.ExpFilePath, "\")))
    End If
    Service = sService
End Property
Friend Property Let ExpFileExtension(ByVal s As String)
    sExpFileExtension = s
    sExpFileFullName = Me.ExpFilePath & "\" & Me.CompName & s
End Property

Friend Property Get Changed( _
             Optional ByVal check_again As Boolean = False) As Boolean
' --------------------------------------------------------------------
' returns TRUE when the raw's regular Export File differs from the
' regular Export File of the clone which indicates that the code of
' the raw had changed.
' --------------------------------------------------------------------
    Const PROC = "Changed-Get"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If dctChanges Is Nothing Or check_again Then
        Set dctChanges = _
        mFile.Differs(fd_file1:=fso.GetFile(sCloneExpFileFullName) _
                    , fd_file2:=fso.GetFile(Me.ExpFileFullName) _
                    , fd_ignore_empty_records:=True _
                    , fd_compare:=vbTextCompare _
                    , fd_stop_after:=1 _
                     )

    End If
    Changed = dctChanges.Count <> 0
    If Changed Then
        cLog.Action = "Raw changed! (raw export differs from clone export)"
    Else
        cLog.Action = "Raw not changed! (raw export and clone export are identical)"
    End If
        
xt: Set fso = Nothing
    Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Let CloneExpFileFullName(ByVal s As String):    sCloneExpFileFullName = s:                      End Property

Friend Property Get CodeAsOfDate() As String:               CodeAsOfDate = flExpFile.DateLastModified:  End Property

Friend Property Get CompName() As String:                   CompName = sCompName:                       End Property

Friend Property Let CompName(ByVal s As String):            sCompName = s:                              End Property

Friend Property Get DsplyAllChanges() As Long
' -------------------------------------------
' Display of all differences between the
' clone's and the raw's Export File.
' -------------------------------------------
    Const PROC = "DisplayAllChanges"
    
    On Error GoTo eh
    Dim sExpFileTemp    As String
    Dim wb              As Workbook
    Dim cComp           As New clsComp
    Dim fso             As New FileSystemObject
    Dim sTmpFolder     As String
    Dim flExpTemp       As File
    Dim sMsg            As tMsg
    
    If Me.Changed Then
        mFile.Compare fc_file_left:=sCloneExpFileFullName _
                    , fc_file_right:=sExpFileFullName _
                    , fc_left_title:="'Clone' component's export file '" & sCloneExpFileFullName & "'" _
                    , fc_right_title:="'Raw' component's export file '" & sExpFilePath & "'"
    Else
        sMsg.Section(1).sText = "The Export File of the component '" & Me.CompName & "' (" & sCloneExpFileFullName & ") " & _
                       "and the raw's Export File (" & sExpFilePath & ") are identical. The clone component " & _
                       "is thus up-to-date."
        sMsg.Section(2).sLabel = "Please note!"
        sMsg.Section(2).sText = "Differences in empty code lines or upper/lower case text do not constitute a relevant difference."
        mMsg.Dsply msg_title:="Display of all raw changes/differences failed!" _
                 , msg:=sMsg
    End If
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Get ExpFile() As File
    If flExpFile Is Nothing Then
        With New FileSystemObject
            Set flExpFile = .GetFile(sExpFileFullName)
        End With
    End If
    Set ExpFile = flExpFile
End Property

Friend Property Let ExpFile(ByVal fl As File)
    
    If fl Is Nothing Then
        With New FileSystemObject
            If .FileExists(Me.ExpFileFullName) Then Set fl = .GetFile(Me.ExpFileFullName)
        End With
    End If
    Set flExpFile = fl

End Property

Friend Property Get ExpFileFullName() As String:        ExpFileFullName = sExpFileFullName:     End Property

Friend Property Get ExpFilePath() As String:            ExpFilePath = sExpFilePath:             End Property

Friend Property Get HostFullName() As String:           HostFullName = sHostFullName:           End Property

Friend Property Let HostFullName(ByVal s As String)
    sHostFullName = s
    If sExpFilePath = vbNullString Then
        With New FileSystemObject
            sExpFilePath = .GetParentFolderName(s)
        End With
    End If
    sService = Split(sExpFilePath, "\")(UBound(Split(sExpFilePath, "\")))
End Property

Friend Property Get RawHost() As Workbook:              Set RawHost = wbRawHost:        End Property

Friend Property Let RawHost(ByVal wb As Workbook):      Set wbRawHost = wb:             End Property

Friend Property Get RawType() As vbext_ComponentType:   RawType = vbcRaw.Type:          End Property

Friend Property Get RawVbc() As VBComponent:            Set RawVbc = vbcRaw:            End Property

Friend Property Let RawVbc(ByVal vbc As VBComponent):   Set vbcRaw = vbc:               End Property

Friend Property Get WrkbkBaseName() As String
    With New FileSystemObject
        WrkbkBaseName = .GetBaseName(sHostFullName)
    End With
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsRaw" & "." & sProc
End Function

