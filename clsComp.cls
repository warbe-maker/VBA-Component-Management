VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------------------------------------------------------
' Class Module clsComp
'       Represents a VB-Project's Component with extended (Management) properties and methods.
'       The Component may be of either of the following kinds:
'       - a used Component which may also be a used Common Component
'       - the original/raw of a Common Component, i.e. a Component of which the original is
'         hosted in a dedicated Common Component Workbookorigin code of a Common Component
'
' Properties:
' - ExpFile      R/W File object created by menas of the BackUpCode method
' - HostedRaws  W   String with named Components hosted in the Workbook
' - IsRawHostedElsewhere           R   TRUE when a Component is regarded a Common one
' - CodeChanged     R   TRUE when the code of the CodeModule differs from the code
'                       backed up in the corresponding ExportFile
' - KindOfChange    R
' - KindOfComp      R
'
' Methods:
' - BackUpCode
'
'
' -----------------------------------------------------------------------------------
Private sWrkbkBaseName      As String       ' Common Component host Workbook's base name
Private sWrkbkFullName      As String       ' Common Component host Workbook's full name
Private sCompName           As String       ' The class module's global component name
Private wb                  As Workbook     ' The class' Workbook
Private vbc                 As VBComponent  ' VBComponent managed by this Class Module
Private dtUpdatedAsOf       As Date
Private flExpFile           As FILE
Private foExport            As Folder
Private sExpFilePath        As String       ' Full filespec for the Export File created by the method BackUpCode
Private sExpPath            As String
Private lCompMaxLen         As Long
Private lKindOfComp         As enKindOfComp
Private sExtension          As String
Private dctChanges          As Dictionary   ' Dictionary of code changes
Private sTempExpFilePath    As String
Private sTempFolder         As String

Friend Property Get CompName() As String:           CompName = sCompName:   End Property

Friend Property Let CompName(ByVal s As String)
' ---------------------------------------------
' Provides the properties:
' - CompName (sCompName)
' - VBComp (vbc)
' - Extension (sExtension),
' - ExpFilePath (sExpFilePath)
' ---------------------------------------------

    sCompName = s
    If Not wb Is Nothing Then
        Set vbc = wb.VBProject.VBComponents(sCompName)
        Select Case vbc.Type
            Case vbext_ct_StdModule:    sExtension = ".bas"
            Case vbext_ct_ClassModule:  sExtension = ".cls"
            Case vbext_ct_MSForm:       sExtension = ".frm"
            Case vbext_ct_Document:     sExtension = ".cls"
        End Select
        sExpFilePath = sExpPath & "\" & sCompName & sExtension
    End If

End Property

Friend Property Get ExpFile() As FILE
    
    Dim fso As New FileSystemObject
    
    If flExpFile Is Nothing Then
        If fso.FileExists(sExpFilePath) Then
            Set flExpFile = GetFile(sExpFilePath)
        Else
            '~~ The component apparently has yet no never been exported
            If Not Me.VBComp Is Nothing Then
                Me.VBComp.Export sExpFilePath
                Set flExpFile = GetFile(sExpFilePath)
            End If
        End If
    End If
    
    Set ExpFile = flExpFile
    Set fso = Nothing
    
End Property

Friend Property Let ExpFile(ByVal fl As FILE):      Set flExpFile = fl:     End Property

Private Property Get ExpFileExtension() As String
    Select Case vbc.Type
        Case vbext_ct_StdModule:    ExpFileExtension = ".bas"
        Case vbext_ct_ClassModule:  ExpFileExtension = ".cls"
        Case vbext_ct_MSForm:       ExpFileExtension = ".frm"
        Case vbext_ct_Document:     ExpFileExtension = ".cls"
    End Select
End Property

Friend Property Get ExpFilePath() As String:    ExpFilePath = sExpFilePath: End Property

Friend Property Get ExpFolder() As Folder:      Set ExpFolder = foExport:   End Property

Friend Property Get Extension() As String:      Extension = sExtension:     End Property

Friend Property Get DsplyAllChanges() As Long
' -------------------------------------------
' Display of all differences between the
' clone's and the raw's Export File.
' -------------------------------------------
    Const PROC = "DisplayAllChanges"
    
    On Error GoTo eh
    Dim sMsg As tMsg
    
    If Me.Changed Then
        mFile.Compare file_left_full_name:=sTempExpFilePath _
                    , file_right_full_name:=sExpFilePath _
                    , file_left_title:="'Clone' component's export file '" & sTempExpFilePath & "'" _
                    , file_right_title:="'Raw' component's export file '" & sExpFilePath & "'"
    Else
        sMsg.section(1).sText = "The Export File of the component '" & Me.CompName & "' reflects the current code. " & _
                                "Thus there are no code changes to be displayed."
        sMsg.section(2).sLabel = "Please note!"
        sMsg.section(2).sText = "Differences in empty code lines or upper/lower case text do not constitute a relevant code change."
        mMsg.Dsply msg_title:="Display of code changes failed!" _
                 , msg:=sMsg
    End If
    
xt: Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Get KindOfComp() As enKindOfComp
    Const PROC = "KindOfComp"
    
    On Error GoTo eh
    
    If lKindOfComp = enKindOfComp.enUnknown Then
        If mHostedRaws.Exists(sCompName) Then
            If mHostedRaws.HostFullName(comp_name:=sCompName) = Me.Wrkbk.FullName _
            Then lKindOfComp = enHostedRaw _
            Else lKindOfComp = enRawClone
        ElseIf lKindOfComp = enKindOfComp.enUnknown Then
            lKindOfComp = enInternal
        End If
        cLog.Action = "Kind of Component = " & KoCStrng(lKindOfComp)
    End If
    
xt: KindOfComp = lKindOfComp
    Exit Property

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property


Private Function KoCStrng(ByVal koc As enKindOfComp) As String
    Select Case koc
        Case enKindOfComp.enUnknown:        KoCStrng = "Unknown"
        Case enKindOfComp.enInternal:       KoCStrng = "Internal (neither raw nor clone)"
        Case enKindOfComp.enRawClone:       KoCStrng = "Raw Clone"
        Case enKindOfComp.enHostedRaw:      KoCStrng = "Hosted Raw"
    End Select
End Function

Private Property Get ReplaceRawWithCloneConfirmed() As Boolean
' -------------------------------------------------------------
' Returns TRUE when the "reverse" code update is confirmed.
' --------------------------------------------------------------
    Const PROC = "ReplaceRawWithCloneConfirmed"
    
    On Error GoTo eh
    Dim sReplyUpdateOrigin      As String
    Dim sReplyDsplyDiff         As String
    Dim sReplyIgnore            As String
    Dim sMsg                    As tMsg
    Dim cllButtons              As Collection
    Dim sTitle                  As String
    Dim sReply                  As String
    
    '~~ Prepare communication message and reply option
    sReplyUpdateOrigin = "Update the origin code:" & vbLf & vbLf & _
                         "Replace (remove and re-import) the component " & vbLf & _
                         "" & _
                         "in Workbook " & vbLf & _
                         "" & vbLf & _
                         " with this component's Export File. I.e. the changes" & vbLf & _
                         "made will become ""permanent"""
    sReplyDsplyDiff = "Display the difference:" & vbLf & vbLf & _
                      "The installed ""WinMerge"" is used. Empty code" & vbLf & _
                      "lines are ignored and the compare is case-insensitive."
    sReplyIgnore = "Ignore the code change:" & vbLf & vbLf & _
                   "The change will be overwritten by the" & vbLf & _
                   "next open of this Workbook."
    Set cllButtons = mMsg.Buttons(sReplyUpdateOrigin, sReplyDsplyDiff, sReplyIgnore)
    
    '~~ Offer options
    sTitle = "A ""Remotely hosted raw's clone has been modified in this VB-Project!"
    With sMsg
        .section(1).sLabel = "About this incident:"
        .section(1).sText = "The module '" & cComp.CompName & "' is a clone of the raw component hosted in Workbook '" & cRaw.WrkbkBaseName & "' " & _
                            "This clone component's code has usually has been synched with the remote raw componet's code. " & _
                            "The remote raw hasn't been changed but the cloned component's code has."
        .section(2).sLabel = "About the choices:"
        .section(2).sText = "When the code modification is ignored it will be regarded temporarily only and overwrtitten with the next open of this Workbook. " & _
                            "When the remote The module, updated when this Workbook has opened, is no longer identical with its origin. " & _
                            "I.e. the component's code has been changed in this Workbook using it instead of in the " & _
                            "Workbook the origin is ""hosted""."
        .section(3).sText = "Take your decission or postpone it after having displayed the difference."
    End With
        
    sReply = sReplyDsplyDiff
    Do
        sReply = mMsg.Dsply(msg_title:=sTitle, _
                            msg:=sMsg, _
                            msg_buttons:=cllButtons _
                           )
        Select Case sReply
            Case sReplyUpdateOrigin
                ReplaceRawWithCloneConfirmed = True
                Exit Do
            Case sReplyDsplyDiff
                mFile.Compare file_left_full_name:=cComp.ExpFilePath _
                          , file_right_full_name:=cRaw.ExpFilePath _
                          , file_left_title:="The cloned raw's current code in Workbook/VBProject " & cComp.WrkbkBaseName & " (" & cComp.ExpFilePath & ")" _
                          , file_right_title:="The remote raw's current code in Workbook/VBProject " & cRaw.WrkbkBaseName & " (" & cRaw.ExpFilePath & ")"
                ' The question is re-displayed for a decison
            Case sReplyIgnore
                ReplaceRawWithCloneConfirmed = False
                Exit Do
        End Select
    Loop
                                            
xt: Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Get VBComp() As VBComponent
    Const PROC = "VBComp_Get"
    
    On Error GoTo eh
    
    If vbc Is Nothing Then
        If Not mCompMan.CompExists(ce_wb:=wb, ce_comp_name:=sCompName) _
        Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "A Component named '" & sCompName & "' does not exist in Workbook '" & wb.name & "'!"
        Set vbc = wb.VBProject.VBComponents(sCompName)
    End If
    Set VBComp = vbc

xt: Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Let VBComp(ByVal vb_comp As VBComponent):       Set vbc = vb_comp:                      End Property

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
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Friend Property Let Wrkbk(ByVal wbk As Workbook)
' ----------------------------------------------
' Provides the properties:
' - Wrkbk (wb),
' - WrkbFullName (sWrkbkFullName)
' - WrkbkBaseName (sWrkbkBaseName)
' ---------------------------------------------
    
    Dim fso As New FileSystemObject
    
    Set wb = wbk
    sWrkbkFullName = wb.FullName
    sWrkbkBaseName = fso.GetBaseName(sWrkbkFullName)
    Set foExport = fso.GetFolder(Me.ExpPath)
    For Each vbc In wb.VBProject.VBComponents
        lCompMaxLen = mBasic.Max(lCompMaxLen, Len(vbc.name))
    Next vbc
    Set fso = Nothing
    
End Property

Friend Property Get WrkbkBaseName() As String:                 WrkbkBaseName = BaseName(sWrkbkFullName):        End Property

Friend Property Get WrkbkFullName() As String:                 WrkbkFullName = sWrkbkFullName:                  End Property

Friend Property Let WrkbkFullName(ByVal s As String):          sWrkbkFullName = s:                              End Property

Private Function BaseName(ByVal s As String) As String
    Dim fso As New FileSystemObject
    With fso:  BaseName = .GetBaseName(s): End With
    Set fso = Nothing
End Function

Friend Property Get Changed() As Boolean
' --------------------------------------
' returns TRUE when the temporary Export
' File differs from the regular Export
' File which indicates that the code had
' changed.
' --------------------------------------
    Dim fso         As New FileSystemObject

    If dctChanges Is Nothing Then
        With fso
            sTempFolder = sExpPath & "\Temp"
            If Not .FolderExists(sTempFolder) Then .CreateFolder sTempFolder
            sTempExpFilePath = sTempFolder & "\" & vbc.name & sExtension
            vbc.Export sTempExpFilePath
        End With
    
        Set dctChanges = _
        mFile.Differs(dif_file1:=fso.GetFile(sTempExpFilePath) _
                    , dif_file2:=cComp.ExpFile _
                    , dif_ignore_empty_records:=True _
                    , dif_compare:=vbTextCompare _
                    , dif_stop_after:=1 _
                     )

    End If
    Changed = dctChanges.Count <> 0
    Set fso = Nothing
    
End Property

Private Function DebugComp() As String
    If lMaxCompLength > 0 _
    Then DebugComp = sCompName & " " & String(lMaxCompLength - Len(sCompName), ".") & ": " _
    Else DebugComp = sCompName & " : "
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsComp" & "." & sProc
End Function

Public Function Exists(ByVal sComp As String) As Boolean
    Dim vbc As VBComponent

    On Error Resume Next
    Set vbc = Me.Wrkbk.VBProject.VBComponents(sComp)
    Exists = Err.Number = 0
    Set vbc = Nothing

End Function

Friend Property Get ExpPath() As String
' -----------------------------------------
' Returns the path for the code backup
' export file  of the Component which
' defaults to the Workbook's path with a
' Workbook's base name folder.
' Will be created if not existent.
' --------------------------------------
    Const PROC = "ExpPath_Get"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim sPath   As String
    
    If sExpPath = vbNullString Then
        With fso
            If Right(.GetParentFolderName(wb.FullName), Len(sWrkbkBaseName)) <> sWrkbkBaseName Then
               sPath = .GetParentFolderName(wb.FullName) & "\" & sWrkbkBaseName
                '~~ Make sure the returned export path specifies an existing folder
                If Not .FolderExists(sPath) Then
                    .CreateFolder (sPath)
                End If
            Else
                sPath = .GetParentFolderName(wb.FullName)
            End If
            sExpPath = sPath
        End With
    End If
    
xt: ExpPath = sExpPath
    Set fso = Nothing
    Exit Property
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Property

Public Function ExpFileToArray(ByVal fl As FILE) As String()
' ---------------------------------------------------------------------
' Returns the content of the file (vFile) which may be provided as file
' object or full file name string as array, considering any kind of
' line break characters.
' ---------------------------------------------------------------------
    Const PROC = "ExpFileToArray"
    Const ATTRIB_VB = "Attribute VB_"

    On Error GoTo eh
    Dim a           As Variant
    Dim i           As Long
    
    a = mFile.Arry(fl.Path)
    '~~ Remove the extra module information lines
    For i = UBound(a) To LBound(a) Step -1
        '~~ Find last Exported module info
        If Left(a(i), Len(ATTRIB_VB)) = ATTRIB_VB Then
            Exit For
        End If
    Next i
    If i >= LBound(a) Then mBasic.ArrayRemoveItems a, Element:=1, NoOfElements:=i + 1
    
    '~~ Remove any leading or trailing empty items
    If mBasic.ArrayIsAllocated(a) Then
        mBasic.ArrayTrimm a
        If a(UBound(a)) = vbNullString Then Stop
    End If
    
     ' remove leading and trailing empty items
    mBasic.ArrayTrimm a
    ExpFileToArray = a

xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Public Function ExpFileToDict(ByVal fl As FILE) As Dictionary
' --------------------------------------------------------------
' Returns the content of the export file (fl) as Dictionary.
' All lines preceeding "Attribute " lines and "Attribute " lines
' are excempted.
' --------------------------------------------------------------
    Const PROC = "ExpFileToArray"
    Const ATTRIB_LINE = "Attribute "

    On Error GoTo eh
    Dim i               As Long
    Dim dct             As Dictionary
    Dim v               As Variant
    Dim cllRemove       As New Collection
    Dim sLine           As String
    Dim cll             As Collection
    
    Set cll = New Collection
    Set dct = mFile.ToDict(fl)
    
    '~~ Remove the extra module information lines
    For Each v In dct
        If Left(sLine, Len(ATTRIB_LINE)) = ATTRIB_LINE _
        And Left(dct(v), Len(ATTRIB_LINE)) <> ATTRIB_LINE _
        Then Exit For
        cllRemove.Add v
        sLine = dct(v)
    Next v
    For Each v In cllRemove
        dct.Remove v
    Next v
    Set cllRemove = Nothing
    
    '~~ Remove any leading empty items/lines
    Set cllRemove = New Collection
    For Each v In dct
        If dct(v) <> vbNullString Then
            Exit For
        End If
        Debug.Print "Collect fo removeal " & i & " item with content >>" & dct.Item(i) & "<<"
        cllRemove.Add i:    i = i + 1
        sLine = dct(v)
    Next v
    While Len(dct.Items(dct.Count - 1)) < 2
        dct.Remove dct.Keys((dct.Count - 1))
        Debug.Print "Leading empty item/line removed"
    Wend
    
    '~~ Remove any trailing empty items/lines
    For i = 1 To cllRemove.Count
        dct.Remove cllRemove.Item(i)
        Debug.Print "Trailing empty item/line removed"
    Next i
    
    Set ExpFileToDict = dct

xt: Set cll = Nothing
    Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function GetFile(ByVal sFile As String) As FILE
    Dim fso As New FileSystemObject
    With fso:  Set GetFile = .GetFile(sFile):  End With
    Set fso = Nothing
End Function

Public Sub ReplaceRawWithClone( _
                         ByVal raw_remote_comp_name As String, _
                         ByRef rr_log As clsLog)
' --------------------------------------------------------------
' Replace the raw component's code with the Export File of the
' clone component.
' --------------------------------------------------------------
    Const PROC = "ReplaceRawWithClone"

    On Error GoTo eh
    Dim wbRawHost   As Workbook
        
    Application.EnableEvents = False ' prevent the UpdateRawClones service
    Set wbRawHost = mCompMan.WbkGetOpen(cRaw.HostFullName) ' Wrkbk is the Workbook hosting the raw Component
    Application.EnableEvents = True
    
    mRenew.ByImport rn_wb:=wbRawHost _
                  , rn_comp_name:=sCompName _
                  , rn_exp_file_full_name:=sExpFilePath
    cLog.Action = "Remote raw (hosted in '" & wbRawHost.name & "' replaced by (re-)import of the clone's export file '" & sExpFilePath & "'"

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub ReplaceRawWithCloneWhenConfirmed( _
                                      ByRef rwu_updated As Boolean, _
                                      ByRef rwu_log As clsLog)
' -------------------------------------------------------------------
' Provided confirmed by the user, replaces the raw code in the
'  hosting Workbook with the cloned raw component's code in its
' Export File. Code replacement for class modules, satandard modules
' and UserForms is done via a rename, re-import, remove renamed
' procedure, for a data module the code is replaced line by line
' because a (re-)import is impossible for such a module.
' -------------------------------------------------------------------
    Const PROC = "ReplaceRawWithCloneWhenConfirmed"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    If ReplaceRawWithCloneConfirmed Then
            '~~ Get the origin host Workbook open
        With cRaw
            .HostFullName = mHostedRaws.HostFullName(comp_name:=cComp.CompName)
            .ExpFile = fso.GetFile(mHostedRaws.ExpFilePath(comp_name:=cComp.CompName))
            .ExpFilePath = .ExpFile.Path
            .RawHost = mCompMan.WbkGetOpen(.HostFullName)
            .RawVbc = .RawHost.VBProject.VBComponents(Index:=cComp.CompName)
            Select Case .RawType
                Case vbext_ct_Document
                    '~~ A data module cannot be replaced like the other ones. Thus only its code
                    '~~ is synchronized - disregarding any consequences this may have when the
                    '~~ structure and design of the underlying Workbook or Worksheet has changed.
                    cComp.SynchRawWithCloneByCodeLines
                Case vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_MSForm
                    Me.ReplaceRawWithClone raw_remote_comp_name:=.CompName, rr_log:=rwu_log
            End Select
            rwu_updated = True
        End With ' cRaw
    End If

xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SynchCloneWithRawByCodeLines()
' ---------------------------------------
' Synchronizes cloned raw component's
' code with the remote raws component's
' code in its Export File.
' ---------------------------------------
    Const PROC = "SynchCloneWithRawByCodeLines"

    On Error GoTo eh
    Dim iLine   As Long
    Dim v       As Variant
    Dim vbcm    As CodeModule
    Dim fso     As New FileSystemObject
    
    '~~ Remove all lines from the cloned raw component
    Set vbcm = vbc.CodeModule
    With vbcm
        .DeleteLines 1, .CountOfLines
    
        '~~ Insert/copy all code lines from the source to the target component
        iLine = 0
        For Each v In ExpFileToArray(cRaw.ExpFile)
            iLine = iLine + 1
            .InsertLines iLine, v
        Next v
    End With
    
    Debug.Print DebugComp & Now() & " Code of the cloned raw component's CodeModule synchronized with code the remote raw component's Export File '" & cRaw.ExpFilePath & "'."
                
xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub SynchRawWithCloneByCodeLines()
' ---------------------------------------
' Synchronizes a Common Component's raw
' with the Export File of the clone
' Component.
' ---------------------------------------
    Const PROC = "SynchRawWithCloneByCodeLines"

    On Error GoTo eh
    Dim iLine   As Long
    Dim v       As Variant
    Dim vbcm    As CodeModule
    Dim fso     As New FileSystemObject
                
    Debug.Print DebugComp & Now() & " About to synch the remote raw component's code with the code in the cloned raw's code in the Export File '" & cComp.ExpFilePath & "'."
    '~~ Remove all lines from the target component
    Set vbcm = cRaw.RawVbc.CodeModule
    With vbcm
        .DeleteLines 1, .CountOfLines
        '~~ Insert all code lines from the cloned raw component's Export File into the remote raws CodeModule
        iLine = 0
        For Each v In ExpFileToArray(cComp.ExpFile)
            iLine = iLine + 1
            vbcm.InsertLines iLine, v
        Next v
    End With
    
    Debug.Print DebugComp & Now() & " Code of remote raw component synchronized with the code in the cloned raw's Export File '" & cComp.ExpFilePath & "'."
                
xt: Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Function TypeString() As String
' --------------------------------------------------------------------------
' Returns the component type derived from sComp.
' --------------------------------------------------------------------------
    Const PROC = "TypeString"
    
    On Error GoTo eh
        
    Select Case vbc.Type
        Case vbext_ct_ActiveXDesigner:  TypeString = "ActiveX Designer"
        Case vbext_ct_ClassModule:      TypeString = "Class Module"
        Case vbext_ct_Document:         TypeString = "Document Module"
        Case vbext_ct_MSForm:           TypeString = "UserForm"
        Case vbext_ct_StdModule:        TypeString = "Code Module"
        Case Else:                      TypeString = "Type '" & CStr(vbc.Type) & "' unknown"
    End Select
    
xt: Exit Function
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Function

Private Function UserFormSizeIsDifferent( _
                                   ByVal sid_file_1 As FILE, _
                                   ByVal sid_file_2 As FILE) As Boolean
    Dim sFrxFile1Path  As String
    Dim sFrxFile2Path  As String
    Dim fso             As New FileSystemObject
    
    sFrxFile1Path = Replace$(sid_file_1.Path, ".frm", ".frx")
    sFrxFile2Path = Replace$(sid_file_2.Path, ".frm", ".frx")
    
    With fso
        If sid_file_1.Size <> sid_file_2.Size Then
            UserFormSizeIsDifferent = True
        ElseIf .GetFile(sFrxFile1Path).Size <> .GetFile(sFrxFile2Path).Size Then
            UserFormSizeIsDifferent = True
        End If
    End With
    Set fso = Nothing
    
End Function

Private Sub Class_Terminate()
    Set dctChanges = Nothing
    Set wb = Nothing
    Set vbc = Nothing
    With New FileSystemObject
        If .FolderExists(sTempFolder) Then .DeleteFolder sTempFolder
    End With
    Set flExpFile = Nothing
End Sub
