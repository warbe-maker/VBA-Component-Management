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
' Class Module clsRaw
' Provides Properties and services for a Raw component of any
' type. I.e. a component a certain Workbook claims the raw
' version or the last developed version respectively.
'
' Friend Properties:
' - Changed
' - CloneExpFileFullName
' - CompName
' - CodeLines
' - Exists
' - ExistsBySheetName
' - ExpFile
' - ExpFileExtension
' - ExpFileFullName
' - ExpFilePath
' - HostFullName
' - RawHost
' - RawType
' - RawVbc
' - Service
' - SheetName
' - TypeString
' - Wrkbk
' - WrkbkBaseName
'
' Services:
' - DsplyAllChanges
' - RemoteSheetExists
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
Private wb                      As Workbook
Private sWrkbkBaseName          As String       ' Common Component host Workbook's base name
Private sWrkbkFullName          As String       ' Common Component host Workbook's full name
Private vbc                     As VBComponent  ' VBComponent managed by this Class Module
Private dctType                 As Dictionary
Private lMaxLenComp             As Long
Private foExport                As Folder

Private Sub Class_Initialize()
    If dctType Is Nothing Then Set dctType = New Dictionary
    dctType.Add vbext_ct_ActiveXDesigner, "ActiveX Designer"
    dctType.Add vbext_ct_ClassModule, "Class Module"
    dctType.Add vbext_ct_Document, "Document Module"
    dctType.Add vbext_ct_MSForm, "UserForm"
    dctType.Add vbext_ct_StdModule, "Standard Module"
End Sub

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
            
xt: Set fso = Nothing
    Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Let CloneExpFileFullName(ByVal s As String):    sCloneExpFileFullName = s:                  End Property

Friend Property Get CodeLines() As Dictionary
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

Friend Property Let CompName(ByVal s As String)
' -------------------------------------------------------------
' Provides the properties: - CompName (sCompName)
'                          - ExpFileFullName (sExpFileFullName)
'                          - VBComp          (vbc)
' -------------------------------------------------------------
    sCompName = s
    If Not wb Is Nothing Then
        If Me.Exists Then
            Set vbc = wb.VBProject.VBComponents(s)
            sExpFileFullName = Me.ExpFilePath & "\" & sCompName & Me.ExpFileExtension
        End If
    End If
End Property

Public Sub DsplyAllChanges()
' -------------------------------------------
' Display of all differences between the
' clone's and the raw's Export-File. Uses WinMerge!
' -------------------------------------------
    mFile.Compare fc_file_left:=sCloneExpFileFullName _
                , fc_file_right:=sExpFileFullName _
                , fc_left_title:="'Clone' component's export file '" & sCloneExpFileFullName & "'" _
                , fc_right_title:="'Raw' component's export file '" & sExpFilePath & "'"
End Sub

Friend Property Get Exists(Optional ByRef ex_wb As Workbook = Nothing) As Boolean
    Dim v As VBComponent
    If Not ex_wb Is Nothing Then
        On Error Resume Next
        Set v = ex_wb.VBProject.VBComponents(sCompName)
        Exists = Not v Is Nothing
    ElseIf vbc Is Nothing Then
        On Error Resume Next
        Set vbc = wb.VBProject.VBComponents(sCompName)
        Exists = Not vbc Is Nothing
    Else
        Exists = Not vbc Is Nothing
    End If
End Property

Friend Property Get ExistsBySheetName(ByVal eb_name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Me.Wrkbk.Worksheets
        ExistsBySheetName = ws.Name = eb_name
        If ExistsBySheetName Then Exit For
    Next ws
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

Friend Property Get ExpFileExtension() As String
    
    If sExpFileExtension = vbNullString Then
        If vbc Is Nothing Then Set vbc = wb.VBProject.VBComponents(sCompName)
        With Me
            Select Case vbc.Type
                Case vbext_ct_StdModule:    sExpFileExtension = ".bas"
                Case vbext_ct_ClassModule:  sExpFileExtension = ".cls"
                Case vbext_ct_MSForm:       sExpFileExtension = ".frm"
                Case vbext_ct_Document:     sExpFileExtension = ".cls"
            End Select
        End With
    End If
    ExpFileExtension = sExpFileExtension

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

Friend Property Get TypeString() As String
' ----------------------------------------
' Returns the component's type as String.
' ----------------------------------------
    If Not vbc Is Nothing Then
        On Error Resume Next
        TypeString = dctType(vbc.Type)
        If TypeString = vbNullString Then TypeString = "unknown"
    End If
End Property

Friend Property Let TypeString(ByVal s As String):      sTypeString = s:                End Property

Friend Property Get Wrkbk() As Workbook
    Const PROC = "Wrkbk_Get"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If wb Is Nothing Then
        If Not fso.FileExists(sWrkbkFullName) _
        Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "A Workbook '" & sWrkbkFullName & "' does not exist!"
        Set wb = mCompMan.WbkGetOpen(sWrkbkFullName)
    Else
        Set Wrkbk = wb
    End If
    
xt: Set fso = Nothing
    Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Set Wrkbk(ByRef wbk As Workbook)
' ---------------------------------------------------------
' Provides the properties: - Wrkbk         (wb)
'                          - WrkbFullName  (sWrkbkFullName)
'                          - WrkbkBaseName (sWrkbkBaseName)
'                          - ExpFilePath   (sExpFilePath)
'                          - Service       (sService)
'                          - MaxLenComp    (lMaxLenComp)
' ---------------------------------------------------------
    
    Dim fso As New FileSystemObject
    Dim vbc As VBComponent
    
    Set wb = wbk
    sWrkbkFullName = wb.FullName
    sWrkbkBaseName = fso.GetBaseName(sWrkbkFullName)
    sExpFilePath = fso.GetParentFolderName(sWrkbkFullName)
    Set foExport = fso.GetFolder(sExpFilePath)
    sService = Split(wb.Path, "\")(UBound(Split(wb.Path, "\")))
    For Each vbc In wb.VBProject.VBComponents
        lMaxLenComp = mBasic.Max(lMaxLenComp, Len(vbc.Name))
    Next vbc
    Set fso = Nothing
    
End Property

Friend Property Get WrkbkBaseName() As String
    With New FileSystemObject
        WrkbkBaseName = .GetBaseName(sHostFullName)
    End With
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsRaw" & "." & sProc
End Function

Public Function RemoteSheetExists( _
                            ByRef rs_wb As Workbook, _
                   Optional ByRef rs_ws_name As String = vbNullString, _
                   Optional ByRef rs_ws_codename As String = vbNullString, _
                   Optional ByRef rs_ws As Worksheet = Nothing, _
                   Optional ByRef rs_vbc As VBComponent = Nothing) As Boolean
' ---------------------------------------------------------------------------
' When the remote sheet exists in the remote Workbook (rs_wb) either under
' the provided remote name (remote_ws_name) or the provided remote CodeName
' (remote-ws_codename):
' - True plus the remote Worksheet object (rs_ws) and the remote VBComponent
'   (rs_vbc).
' When the remote sheet not exists in the remote Workbook (rs_wb) neither
' under the provided remote name (remote_ws_name) nor under the provided
' remote CodeName (remote-ws_codename):
' - False and neither a remote sheet object nor a remote VBComponent.
' ---------------------------------------------------------------------------
    Dim ws  As Worksheet
    Dim vbc As VBComponent
    
    For Each ws In rs_wb.Worksheets
        If rs_ws_name <> vbNullString And ws.Name = rs_ws_name Then
            Set rs_ws = ws
            If rs_ws_codename = vbNullString Then rs_ws_codename = ws.CodeName
            Exit For
        ElseIf rs_ws_codename <> vbNullString And ws.CodeName = rs_ws_codename Then
            Set rs_ws = ws
            If rs_ws_name = vbNullString Then rs_ws_name = ws.Name
        End If
    Next ws
    
    For Each vbc In rs_wb.VBProject.VBComponents
        If rs_ws_codename <> vbNullString And vbc.Name = rs_ws_codename Then
            Set rs_vbc = vbc
            Exit For
        End If
    Next vbc
   
End Function

