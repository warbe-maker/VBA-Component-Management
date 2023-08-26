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
'
' ----------------------------------------------------------------------------------------
Private sFullName   As String
Private dctNames    As Dictionary
Private fso         As FileSystemObject

Private Sub Class_Initialize()
    Set fso = New FileSystemObject
    Set dctNames = New Dictionary
End Sub

Friend Property Let Names(ByVal n_dct As Dictionary)
    Set dctNames = n_dct
    HskpngValueNames
End Property

Friend Property Let FullName(ByVal s As String)
    sFullName = s
    If Not fso.FileExists(sFullName) Then
        fso.CreateTextFile sFullName
    End If
End Property

Public Property Get Name(Optional ByVal s As String) As String: Name = Split(dctNames(s), ":")(0):   End Property

Private Sub Class_Terminate()
    Set fso = Nothing
    Set dctNames = Nothing
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsPP" & "." & sProc
End Function

Friend Property Get Value(Optional ByVal pp_section As String, _
                          Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file CompManDatFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFso.PPvalue(pp_file:=sFullName _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Function Components() As Dictionary
    Set Components = mFso.PPsectionNames(sFullName)
End Function

Friend Property Let Value(Optional ByVal pp_section As String, _
                          Optional ByVal pp_value_name As String, _
                                   ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (pp_value) under the name (pp_value_name) into the
' CompManDatFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
'    Debug.Print "Write Value-Name: " & pp_value_name & " = Value: " & pp_value & " into Priv-Prof-File: " & sFullName
    mFso.PPvalue(pp_file:=sFullName _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Sub HskpngValueNames()
' ----------------------------------------------------------------------------
' Renames all value names with a syntax <new>:<old> in all sections,
' Removes all value names with a syntax :<old>
' When at least one housekeeping action had been performed the function
' returns TRUE.
' ----------------------------------------------------------------------------
    Const PROC = "HskpngValueNames"
    
    On Error GoTo eh
    Dim dctSects    As Dictionary
    Dim i           As Long
    Dim sNew        As String
    Dim sOld        As String
    Dim v           As Variant
    Dim vName       As Variant
    Dim bDone       As Boolean
    
    mBasic.BoP ErrSrc(PROC)
    
    For i = 0 To dctNames.Count - 1
        vName = Split(dctNames.Items(i), ":")
        If UBound(vName) = 1 Then
            '~~> <new>:<old> or :<remove>
            If vName(0) = vbNullString Then
                '~~ Remove the name in all sections
                If ValueNameRemoveInAllSections(vName(1)) Then
                    bDone = True
                End If
            Else
                '~~> Rename the old name with the new name in all sections
                sNew = vName(0)
                sOld = vName(1)
                If sOld <> vbNullString And sNew <> vbNullString Then
                    If ValueNameRenameInAllSections(sOld, sNew) = True Then
                        bDone = True
                    End If
                End If
            End If
        End If
    Next i

'    If Not bDone _
'    Then Debug.Print "All Value-Names housekeeping in " & sFullName & " done!" _
'    Else Debug.Print "Value-Names housekeeping in " & sFullName & " performed!"

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Friend Property Get ValueSections() As Dictionary
    Set ValueSections = mFso.PPsections(sFullName)
End Property

Private Function ValueNameRemoveInAllSections(ByVal pp_name As String) As Boolean
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    For Each v In ValueSections
        If mFso.PPremoveNames(pp_file:=sFullName _
                            , pp_section:=v _
                            , pp_value_names:=pp_name) Then
            ValueNameRemoveInAllSections = True
        End If
    Next v
    
End Function

Private Function ValueNameRenameInAllSections(ByVal v_old As String, _
                                              ByVal v_new As String) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim v As Variant
    
    For Each v In ValueSections
        If mFso.PPvalueNameRename(v_old, v_new, sFullName, v, False) Then
            ValueNameRenameInAllSections = True
        End If
    Next v
    
End Function

Public Sub Reorg()
    mFso.PPreorg sFullName
End Sub

'Private Function NameExists(ByVal pp_section As String, _
'                            ByVal pp_value_name As String) As Boolean
'    NameExists = mFso.Exists(x_file:=sFullName _
'                           , x_section:=pp_section _
'                           , x_value_name:=pp_value_name)
'End Function

Public Sub ComponentRemove(ByVal r_comp_name As String)
    mFso.PPremoveSections pp_file:=sFullName, pp_sections:=r_comp_name
End Sub

