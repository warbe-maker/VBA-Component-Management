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
Private dctChanges              As Dictionary
Private flExpFile               As File
Private sCloneExpFileFullName   As String
Private sCompName               As String
Private sExpFileExtension       As String
Private sExpFileFullName        As String
Private sExpFilePath            As String
Private sHostFullName           As String
Private vbcRaw                  As VBComponent
Private wbRawHost               As Workbook
Private sService                As String
Private sTypeString             As String

Private Sub Class_Terminate()
    Set dctChanges = Nothing
End Sub

Friend Property Get Changed( _
             Optional ByVal check_again As Boolean = False) As Boolean
' --------------------------------------------------------------------
' Returns TRUE when the raw's regular Export-File differs from the
' clone's regular Export-File which indicates that the code of the raw
' had changed.
' Expception: When the clone's coresponding Export-File does not exist
'             a change is regarded in the sense that the raw's Export-
'             File indicates a new component
' --------------------------------------------------------------------
    Const PROC = "Changed-Get"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If dctChanges Is Nothing Or check_again Then
        If Not fso.FileExists(sCloneExpFileFullName) Then
            Set dctChanges = New Dictionary
            Changed = True
        Else
            Set dctChanges = _
            mFile.Differs(fd_file1:=fso.GetFile(sCloneExpFileFullName) _
                        , fd_file2:=fso.GetFile(Me.ExpFileFullName) _
                        , fd_ignore_empty_records:=True _
                        , fd_compare:=vbTextCompare _
                        , fd_stop_after:=1 _
                         )
            Changed = dctChanges.Count <> 0
        End If
    End If
    
    If Changed _
    Then cLog.Entry = "The corresponding Raw's " & Me.TypeString & " changed! (its Export-File differs from the Clone's Export-File)" _
    Else cLog.Entry = "The corresponding Raw's " & Me.TypeString & " has not changed! (its Export-File is identical with the Clone's Export-File)"
        
xt: Set fso = Nothing
    Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Let CloneExpFileFullName(ByVal s As String):    sCloneExpFileFullName = s:                  End Property

Friend Property Get CodeAsOfDate() As String:                   CodeAsOfDate = flExpFile.DateLastModified:  End Property

Public Property Get CodeLines() As Dictionary
' ----------------------------------------------------------
' Returns the content of an Export-File (Me.ExpFileFullName)
' as array, considering any kind of line break characters.
' ----------------------------------------------------------
    Const PROC = "CodeLines"
    Const ATTRIB_VB = "Attribute VB_"

    On Error GoTo eh
    Dim i   As Long
    Dim j   As Long
    Dim dct As Dictionary
    
    With New FileSystemObject
        Set dct = mFile.Dict(.GetFile(Me.ExpFileFullName))
    End With
    
    '~~ Remove the extra module information lines
    For i = dct.Count - 1 To 0 Step -1
        If Left(dct.Items()(i - 1), Len(ATTRIB_VB)) = ATTRIB_VB Then Exit For
    Next i
    
    For j = 0 To i - 1
        dct.Remove dct.Keys()(0)
    Next j
    
    '~~ Remove any leading or trailing empty items
    If dct.Count = 0 Then GoTo xt
    While dct.Items()(0) = vbNullString
        dct.Remove dct.Keys()(0)
    Wend
    
    ' Remove any trailing empty items
    While dct.Items()(dct.Count - 1) = vbNullString
        dct.Remove dct.Keys()(dct.Count - 1)
    Wend

xt: Set CodeLines = dct
    Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Get CompName() As String:                       CompName = sCompName:                       End Property

Friend Property Let CompName(ByVal s As String):                sCompName = s:                              End Property

Friend Property Get DsplyAllChanges() As Long
' -------------------------------------------
' Display of all differences between the
' clone's and the raw's Export-File.
' -------------------------------------------
    Const PROC = "DisplayAllChanges"
    
    On Error GoTo eh
    Dim sMsg            As tMsg
    
    If Me.Changed Then
        mFile.Compare fc_file_left:=sCloneExpFileFullName _
                    , fc_file_right:=sExpFileFullName _
                    , fc_left_title:="'Clone' component's export file '" & sCloneExpFileFullName & "'" _
                    , fc_right_title:="'Raw' component's export file '" & sExpFilePath & "'"
    Else
        sMsg.Section(1).sText = "The Export Files of the component '" & Me.CompName & "'"
        sMsg.Section(2).sText = "Clone: " & sCloneExpFileFullName & vbLf & _
                                "Raw:   " & sExpFilePath
        sMsg.Section(2).bMonspaced = True
        sMsg.Section(3).sText = "are identical. The clone component is thus up-to-date."
        sMsg.Section(4).sLabel = "Please note!"
        sMsg.Section(4).sText = "Differences in empty code lines or upper/lower case text do not constitute a relevant difference."
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

Friend Property Set ExpFile(ByVal fl As File)
    
    If fl Is Nothing Then
        With New FileSystemObject
            If .FileExists(Me.ExpFileFullName) Then Set fl = .GetFile(Me.ExpFileFullName)
        End With
    End If
    Set flExpFile = fl

End Property

Friend Property Let ExpFileExtension(ByVal s As String)
    sExpFileExtension = s
    sExpFileFullName = Me.ExpFilePath & "\" & Me.CompName & s
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

Friend Property Set RawHost(ByVal wb As Workbook):      Set wbRawHost = wb:             End Property

Friend Property Get RawType() As vbext_ComponentType:   RawType = vbcRaw.Type:          End Property

Friend Property Get RawVbc() As VBComponent:            Set RawVbc = vbcRaw:            End Property

Friend Property Set RawVbc(ByRef vbc As VBComponent):   Set vbcRaw = vbc:               End Property

Friend Property Get Service() As String
    If sService = vbNullString Then
        sService = Split(Me.ExpFilePath, "\")(UBound(Split(Me.ExpFilePath, "\")))
    End If
    Service = sService
End Property

Friend Property Let TypeString(ByVal s As String):      sTypeString = s:                End Property

Friend Property Get TypeString() As String:             TypeString = sTypeString:       End Property

Friend Property Get WrkbkBaseName() As String
    With New FileSystemObject
        WrkbkBaseName = .GetBaseName(sHostFullName)
    End With
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsRaw" & "." & sProc
End Function

