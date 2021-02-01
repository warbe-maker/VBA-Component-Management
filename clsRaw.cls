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
Private wbRawHost           As Workbook
Private sHostFullName       As String
Private sExpFilePath        As String
Private flExpFile           As File
Private sCompName           As String
Private vbcRaw              As VBComponent
Private sCloneExpFilePath   As String
Private dctChanges          As Dictionary

Private Sub Class_Terminate()
    Set dctChanges = Nothing
End Sub

Friend Property Get Changed( _
             Optional ByVal check_again As Boolean = False) As Boolean
' --------------------------------------------------------------------
' returns TRUE when the raw's regular Export File differs from the
' regular Export File of the clone which indicates that the code of
' the raw had changed.
' --------------------------------------------------------------------
    Dim fso         As New FileSystemObject
    Dim sTempFolder As String
    Dim sTempExpFilePath As String
    
    If dctChanges Is Nothing Or check_again Then
        Set dctChanges = _
        mFile.Differs(fd_file1:=fso.GetFile(sCloneExpFilePath) _
                    , fd_file2:=Me.ExpFile _
                    , fd_ignore_empty_records:=True _
                    , fd_compare:=vbTextCompare _
                    , fd_stop_after:=1 _
                     )

    End If
    Changed = dctChanges.Count <> 0
    Set fso = Nothing
    
End Property

Friend Property Let CloneExpFilePath(ByVal s As String):    sCloneExpFilePath = s:                      End Property

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
    Dim sTempFolder     As String
    Dim flExpTemp       As File
    Dim sMsg            As tMsg
    
    If Me.Changed Then
        mFile.Compare fc_file_left:=sCloneExpFilePath _
                    , fc_file_right:=sExpFilePath _
                    , fc_left_title:="'Clone' component's export file '" & sCloneExpFilePath & "'" _
                    , fc_right_title:="'Raw' component's export file '" & sExpFilePath & "'"
    Else
        sMsg.Section(1).sText = "The Export File of the component '" & Me.CompName & "' (" & sCloneExpFilePath & ") " & _
                       "and the raw's Export File (" & sExpFilePath & ") are identical. The clone component " & _
                       "is thus up-to-date."
        sMsg.Section(2).sLabel = "Please note!"
        sMsg.Section(2).sText = "Differences in empty code lines or upper/lower case text do not constitute a relevant difference."
        mMsg.Dsply msg_title:="Display of all raw changes/differences failed!" _
                 , msg:=sMsg
    End If
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Get ExpFile() As File:                  Set ExpFile = flExpFile:                        End Property

Friend Property Let ExpFile(ByVal fl As File):          Set flExpFile = fl:                             End Property

Friend Property Get ExpFilePath() As String:        ExpFilePath = sExpFilePath:             End Property

Friend Property Let ExpFilePath(ByVal s As String): sExpFilePath = s:                           End Property

Friend Property Get HostFullName() As String:           HostFullName = sHostFullName:                   End Property

Friend Property Let HostFullName(ByVal s As String):    sHostFullName = s:                              End Property

Friend Property Get RawHost() As Workbook:              Set RawHost = wbRawHost:                        End Property

Friend Property Let RawHost(ByVal wb As Workbook):      Set wbRawHost = wb:                             End Property

Friend Property Get RawType() As vbext_ComponentType:   RawType = vbcRaw.Type:                          End Property

Friend Property Get RawVbc() As VBComponent:            Set RawVbc = vbcRaw:                            End Property

Friend Property Let RawVbc(ByVal vbc As VBComponent):   Set vbcRaw = vbc:                               End Property

Friend Property Get WrkbkBaseName() As String
    Dim fso As New FileSystemObject
    WrkbkBaseName = fso.GetBaseName(sHostFullName)
    Set fso = Nothing
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsRaw" & "." & sProc
End Function

