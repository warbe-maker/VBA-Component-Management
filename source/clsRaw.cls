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
' Class Module clsRaw:  Provides Properties and services for a Raw component of any
'                       type. I.e. a component a certain Workbook claims the raw
'                       version or the last developed version respectively.
'
' Friend(Public) Properties:
' - Changed
' - CloneExpFileFullName
' - CompName
' - CodeLines
' - Exists
' - ExistsBySheetName
' - ExpFile
' - ExpFileExt
' - ExpFileFullName
' - ExpFilePath
' - HostWrkbkFullName   In contrast to the clsComp Class Module a 'Raw-Component's'
'                       host Workbook is usually not open. The name of the Export-
'                       File is obtained from the .dat files in the 'CompMan-Addin-Folder'
'                       where the data had been registered when the 'Raw-Host's'
'                       'Raw-Component' had been exported with the Workbook_BeforeSave
'                       event.
' - RawHost
' - RawType
' - VBComp
' - Service
' - SheetName
' - TypeString
' - Wrkbk               When the clsRaw represents a 'Synchronization-Source-Workbook' this Workbook
'                       is always open because its Worksheets, Names, Shapes, and References are
'                       obtained
' - WrkbkBaseName
'
' Services:
' - DsplyAllChanges
' - RemoteSheetExists
'
' W. Rauschenberger Berlin, April 2021 (last revised)
' ----------------------------------------------------------------------------------------
Private dctChanges              As Dictionary
Private flExpFile               As File
Private sCloneExpFileFullName   As String
Private sCompName               As String
Private sExpFileExt             As String
Private sExpFileFullName        As String
Private wbRawHost               As Workbook
Private sService                As String
Private sTypeString             As String
Private wb                      As Workbook
Private sWrkbkBaseName          As String       ' Common Component host Workbook's base name
Private sWrkbkFullName          As String       ' Common Component host Workbook's full name
Private vbc                     As VBComponent  ' VBComponent managed by this Class Module
Private dctType                 As Dictionary
Private lMaxLenComp             As Long
Private vbcClone                As VBComponent

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

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

Friend Property Get Changed( _
                      ByRef clone As clsComp, _
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
        '~~ The difference between the raw's and the clone's
        '~~ export file provided only oncce or when explicitely requested
        If Not fso.FileExists(clone.ExpFileFullName) Then
            '~~ The component never had been exported yet or the Export-File has been deleted
            clone.Export
        End If
        Set dctChanges = _
        mExport.FilesDifference(fd_exp_file_1:=fso.GetFile(clone.ExpFileFullName) _
                              , fd_exp_file_2:=fso.GetFile(Me.ExpFileFullName))
    End If
    Changed = dctChanges.Count <> 0
            
xt: Set fso = Nothing
    Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
        If Left(dct.Items()(i), Len(ATTRIB_VB)) = ATTRIB_VB Then Exit For
    Next i
    
    For j = 0 To i
'        Debug.Print dct.Items()(0)
        dct.Remove dct.Keys()(0)
    Next j
    '~~ Remove any leading or trailing empty items
    If dct.Count = 0 Then GoTo xt
    While dct.Items()(0) = vbNullString
        dct.Remove dct.Keys()(0)
    Wend
    
'    Debug.Print dct.Items()(0)
    
    ' Remove any trailing empty items
    While dct.Items()(dct.Count - 1) = vbNullString
        dct.Remove dct.Keys()(dct.Count - 1)
    Wend

xt: Set CodeLines = dct
    Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get CompName() As String:                       CompName = sCompName:                       End Property

Friend Property Let CompName(ByVal s As String)
' -------------------------------------------------------------
' Provides the properties: - CompName (sCompName)
'                          - ExpFileFullName (sExpFileFullName)
'                          - VBComp          (vbc)
' -------------------------------------------------------------
    Dim fso As New FileSystemObject
    sCompName = s
    If Not wb Is Nothing Then
        If Me.Exists Then
            Set vbc = wb.VBProject.VBComponents(s)
            sExpFileFullName = mCompMan.ExpFileFolderPath(Me.HostWrkbkFullName) & "\" & sCompName & Me.ExpFileExt
            If Not fso.FileExists(sExpFileFullName) Then
                MsgBox Title:="An obligatory Export-File is missing!" _
                     , Prompt:="The obligatory Export-File '" & sExpFileFullName & "' for the Component '" & sCompName & "' is missing!" _
                     , Buttons:=vbCritical
                      Stop ' ????
            End If
        End If
    End If
    Set fso = Nothing
End Property

Public Sub DsplyAllChanges()
' ----------------------------------------------------------------------------
' Display of all differences between the clone's and the raw's Export-File.
' Uses WinMerge!
' ----------------------------------------------------------------------------
    mExport.FilesDifferencesDisplay fd_exp_file_left_full_name:=sCloneExpFileFullName _
                                  , fd_exp_file_right_full_name:=sExpFileFullName _
                                  , fd_exp_file_left_title:="'Clone' component's export file '" & sCloneExpFileFullName & "'" _
                                  , fd_exp_file_right_title:="'Raw' component's export file '" & Me.ExpFileFullName & "'"
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

Friend Property Get ExpFileExt() As String
    
    If sExpFileExt = vbNullString Then
        If vbc Is Nothing Then Set vbc = wb.VBProject.VBComponents(sCompName)
        With Me
            Select Case vbc.Type
                Case vbext_ct_StdModule:    sExpFileExt = ".bas"
                Case vbext_ct_ClassModule:  sExpFileExt = ".cls"
                Case vbext_ct_MSForm:       sExpFileExt = ".frm"
                Case vbext_ct_Document:     sExpFileExt = ".cls"
            End Select
        End With
    End If
    ExpFileExt = sExpFileExt

End Property

Friend Property Let ExpFileExt(ByVal s As String)
    sExpFileExt = s
    sExpFileFullName = mCompMan.ExpFileFolderPath(Me.HostWrkbkFullName) & "\" & Me.CompName & s
End Property

Friend Property Get ExpFileFullName() As String:                ExpFileFullName = sExpFileFullName:    End Property

Friend Property Get HostWrkbkFullName() As String:              HostWrkbkFullName = sWrkbkFullName:    End Property

Friend Property Let HostWrkbkFullName(ByVal s As String)
    Dim fso As New FileSystemObject
    sWrkbkFullName = s
    sService = fso.GetFileName(s)
End Property

Friend Property Get VBComp() As VBComponent:                    Set VBComp = vbc:                       End Property

Friend Property Set VBComp(ByRef Comp As VBComponent):          Set vbc = Comp:                         End Property

Friend Property Get Service() As String
    Dim fso As New FileSystemObject
    If sService = vbNullString Then
        sService = fso.GetFileName(Me.HostWrkbkFullName)
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

Friend Property Let TypeString(ByVal s As String):              sTypeString = s:                End Property

Friend Property Get Wrkbk() As Workbook
    Const PROC = "Wrkbk_Get"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If wb Is Nothing Then
        If Not fso.FileExists(sWrkbkFullName) _
        Then Err.Raise AppErr(1), ErrSrc(PROC), "A Workbook '" & sWrkbkFullName & "' does not exist!"
        Set wb = mCompMan.WbkGetOpen(sWrkbkFullName)
    Else
        Set Wrkbk = wb
    End If
    
xt: Set fso = Nothing
    Exit Property
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
    sService = wb.Name
    For Each vbc In wb.VBProject.VBComponents
        lMaxLenComp = mBasic.Max(lMaxLenComp, Len(vbc.Name))
    Next vbc
    Set fso = Nothing
    
End Property

Friend Property Get WrkbkBaseName() As String
    With New FileSystemObject
        WrkbkBaseName = .GetBaseName(sWrkbkFullName)
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

