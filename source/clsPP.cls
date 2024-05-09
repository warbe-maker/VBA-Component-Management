VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------------------------
' Class Module clsPP: CompMan specific Private Profile File services.
' ===================
'
' Public Methods:
' ---------------
' ComponentRemove
' Components
' FileFullName
' Value
' ValueSections
' Name
' Names
' Reorg
'
' Uses: mFso for common Private Profile File services
'
' W. Rauschenberger, Berlin Oct 2023
' ----------------------------------------------------------------------------------------
Private sFileFullName   As String
Private dctNames        As Dictionary
Private FSo             As FileSystemObject
Private PP              As New clsPrivProf

Private Sub Class_Initialize()
    Set FSo = New FileSystemObject
    Set dctNames = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set FSo = Nothing
    Set dctNames = Nothing
End Sub

Friend Property Let FileFullName(ByVal s As String)
    sFileFullName = s
    If Not FSo.FileExists(sFileFullName) Then
        FSo.CreateTextFile sFileFullName
    End If
    PP.FileName = sFileFullName
End Property

Friend Property Get Value(Optional ByVal pp_section As String, _
                          Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file CompManDatFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = PP.Value(pp_value_name, pp_section)
    
xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Let Value(Optional ByVal pp_section As String, _
                          Optional ByVal pp_value_name As String, _
                                   ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (pp_value) with the name (pp_value_name) into the Private-
' Profile-File sFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    PP.Value(pp_value_name, pp_section) = pp_value

xt: Exit Property

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Friend Property Get ValueSections() As Dictionary
    Set ValueSections = PP.Sections
End Property

Public Sub ComponentRemove(ByVal r_comp_name As String)
    PP.SectionRemove r_comp_name
End Sub

Public Function Components() As Dictionary
    Set Components = PP.SectionNames
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsPP" & "." & sProc
End Function

'Private Sub HskpngValueNames()
'' ----------------------------------------------------------------------------
'' The service considers entries in the dctNames Dictionary of which the item
'' has the following syntax: [<new>][:<old>][:<old>]
'' - Renames any <old> value names with an entry's syntax
''   [<new>][:<old>][:<old>]..
''   to <new>,
'' - Removes <old> value names with an entry's syntax
''   :<old>[:<old]...
'' - Returns TRUE when at least one housekeeping action had been performed.
'' ----------------------------------------------------------------------------
'    Const PROC = "HskpngValueNames"
'
'    On Error GoTo eh
'    Dim i           As Long
'    Dim j           As Long
'    Dim sNew        As String
'    Dim sOld        As String
'    Dim vName       As Variant
'    Dim sValueNames As String
'
'    mBasic.BoP ErrSrc(PROC)
'
'    sValueNames = Join(dctNames.Items, ";")
'    If Value(PP.NameHskpngSection, "DoneNamesHskpng") = sValueNames Then GoTo xt
'
'    For i = 0 To dctNames.Count - 1
'        vName = Split(dctNames.Items(i), ":")
'        If UBound(vName) >= 1 Then
'            '~~> One or more old names are to be renamed to a new name<new>:<old> or :<remove>
'            If vName(0) = vbNullString Then
'                For j = 1 To UBound(vName)
'                    '~~ Remove the names in all sections
'                    ValueNameRemoveInAllSections vName(1)
'                Next j
'            Else
'                sNew = vName(0)
'                For j = 1 To UBound(vName)
'                    '~~> Rename in all sections any old name to the new name
'                    sOld = vName(j)
'                    If sOld <> vbNullString And sNew <> vbNullString Then
'                        ValueNameRenameInAllSections sOld, sNew
'                    End If
'                Next j
'            End If
'        End If
'    Next i
'    Value(NamesHousekeepingSection, "DoneNamesHskpng") = sValueNames
'
'xt: mBasic.EoP ErrSrc(PROC)
'    Exit Sub
'
'eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Sub

Public Sub Reorg()
    PP.FileName = sFileFullName
End Sub

Private Function ValueNameRemoveInAllSections(ByVal pp_name As String) As Boolean
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "ValueNameRemoveInAllSections"
    Dim v As Variant
    
    mBasic.BoP ErrSrc(PROC)
    PP.ValueRemove pp_name ' No section = all
    ValueNameRemoveInAllSections = True
    mBasic.EoP ErrSrc(PROC)
    
End Function

Private Function ValueNameRenameInAllSections(ByVal v_old As String, _
                                              ByVal v_new As String) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "ValueNameRenameInAllSections"
    
    Dim v As Variant
    
    mBasic.BoP ErrSrc(PROC)
    For Each v In ValueSections
        PP.ValueNameRename name_value_old:=v_old _
                         , name_value_new:=v_new _
                         , name_sections:=v
        ValueNameRenameInAllSections = True
    Next v
    mBasic.EoP ErrSrc(PROC)
    
End Function

