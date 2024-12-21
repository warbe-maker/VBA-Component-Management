VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommonPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Class Module clsCommonPublic: Management of Common (public) Components in
' ============================= the Common-Components folder, the
' CommComps.dat file and the serviced Workbook's CommComps.dat file.
'
' Public Methods:
' ----------------
' BttnInconsistencyExport          .
' BttnInconsistencySkip            .
' BttnInconsistencyUpdate          .
' Code
' CommCompsDatFileFullName         .
' ComponentRemove
' Exists
' Public
'
' Public Properties:
' ------------------
' CompName                 w Essential throughout all services/methods
' ExpFileFullName
' LastMod...             r/w Properties are provided whenever a used/hosted
'                            Common Component's code has changed, along with
'                            the clsComp.ReleaseComp method.
' PendingReleaseRegister   w Modifications made for a used/hosted Common
'                            Component are registered 'pending release'.
'
' ---------------------------------------------------------------------------
Public CodePublic As clsCode
Private PPFile    As New clsPrivProf
Private dctAll    As Dictionary

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
        
    On Error GoTo eh
    Set PPFile = New clsPrivProf
    
    With PPFile
        .FileName = mEnvironment.CommCompsPublicPrivProfFileFullName
    
        ' Housekeeping names syntax:
        ' [[sect:][sect:]][current-name]<prev-name<prev-name
        ' [[sect:][sect:]]prev-name>prev-name>[current-name]
        ' Rule 1: > or < indicate the history sequence
        ' Rule 2: One ot more sections are optional. When specified, the names history
        '         applies only for those, when omitted for all sections the name any
        '         of the names in the history is used.
        ' Rule 3: The > and the < character point to the predecessor of a name. I.e.
        '         both syntaxes are valid.
        ' Rule 4: The current name is optional. I.e. when > or < is not followed by a
        '         name (the current new name), the life any history name has ended.
        '         I. e. the name is no longer used at all and thus any in use history
        '         name is removed - in the specified or all sections.
        
        .HskpngNames .FileName _
                   , VALUE_NAME_LAST_MOD_AT & "<LastModifiedAtDateTime<RevisionNumber<Last_Modified_AtDateTime" _
                   , VALUE_NAME_LAST_MOD_EXP_FILE_ORIG & "<LastModExpFileFullNameOrigin<LastModExpFileFullName<Last_Modified_ExpFileFullNameOrigin<Last_Modified_ExpFile<<LastModExpFile" _
                   , VALUE_NAME_LAST_MOD_IN & "<LastModifiedInWbkFullName<LastModWbkFullName<Last_Modified_InWbkName<Last_Modified_InWbkFullName" _
                   , VALUE_NAME_LAST_MOD_ON & "<LastModifiedOnMachine<Last_Modified_OnMachine" _
                   , VALUE_NAME_LAST_MOD_BY & "<Last_Modified_ByUser" _
                   , "RawRevisionNumber>" _
                   , "ExportFileExtention>"

        .FileHeader = FileHeader
        .FileFooter = mEnvironment.PrivateProfileFileFooter
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Add(ByVal a_comp As String)
' ----------------------------------------------------------------------------
' Adds a component to the Dictionary of current pubic Common Components.
' ----------------------------------------------------------------------------
    With dctAll
        If Not .Exists(a_comp) Then .Add a_comp, vbNullString
    End With
End Sub

Public Function All() As Dictionary
' ----------------------------------------------------------------------------
' Provides a Dictionary with all current public Common Components.
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    If dctAll Is Nothing Then
        Set dctAll = New Dictionary
        With dctAll
            For Each v In PPFile.Sections
                If Not v Like "_*" And Not v Like "@*" Then
                    .Add v, vbNullString
                End If
            Next v
        End With
    End If
    Set All = dctAll
    
End Function

Private Sub Class_Terminate()
    Set PPFile = Nothing
End Sub

Friend Property Get LastModAt(Optional ByVal l_comp_name As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    LastModAt = Value(VALUE_NAME_LAST_MOD_AT, l_comp_name)
End Property

Friend Property Let LastModKindOfOriginComp(Optional ByVal comp_name As String = vbNullString, _
                                                    ByVal comp_reg_state As enKindOfComp)
    Value(VALUE_NAME_REG_STAT_ORIGIN_COMP, comp_name) = mCompMan.RegState(comp_reg_state)
End Property

Friend Property Get LastModKindOfOriginComp(Optional ByVal comp_name As String) As enKindOfComp
    LastModKindOfOriginComp = mCompMan.RegState(Value(VALUE_NAME_REG_STAT_ORIGIN_COMP, comp_name))
End Property

Friend Property Let LastModAt(Optional ByVal l_comp_name As String = vbNullString, _
                                               ByVal l_comp_rev_no As String)
' ------------------------------------------------------------------------------
' Returns a revision number in the form yy-mm-dd.00. Plus one when an existing
' revision number is provided (comp_rev-no) or the current date with .01.
' ------------------------------------------------------------------------------
    If l_comp_rev_no = vbNullString Then l_comp_rev_no = CompMan.UTC
    Value(VALUE_NAME_LAST_MOD_AT, l_comp_name) = l_comp_rev_no
End Property

Friend Property Get LastModBy(Optional ByVal l_comp_name As String = vbNullString) As String
    LastModBy = Value(VALUE_NAME_LAST_MOD_BY, l_comp_name)
End Property

Friend Property Let LastModBy(Optional ByVal l_comp_name As String = vbNullString, _
                                              ByVal l_machine_name As String)
    Value(VALUE_NAME_LAST_MOD_BY, l_comp_name) = l_machine_name
End Property

Friend Property Get LastModExpFileOrigin(Optional ByVal l_comp_name As String = vbNullString) As String
    LastModExpFileOrigin = Value(VALUE_NAME_LAST_MOD_EXP_FILE_ORIG, l_comp_name)
End Property

Friend Property Let LastModExpFileOrigin(Optional ByVal l_comp_name As String = vbNullString, _
                                            ByVal l_exp_file_full_name As String)
    Value(VALUE_NAME_LAST_MOD_EXP_FILE_ORIG, l_comp_name) = l_exp_file_full_name
End Property

Friend Property Let LastModInWrkbkFullName(Optional ByVal l_comp_name As String = vbNullString, _
                                                  ByVal l_full_name As String)
    Value(VALUE_NAME_LAST_MOD_IN, l_comp_name) = l_full_name
End Property

Friend Property Get LastModInWrkbkFullName(Optional ByVal l_comp_name As String = vbNullString) As String
    LastModInWrkbkFullName = Value(VALUE_NAME_LAST_MOD_IN, l_comp_name)
End Property

Public Function LastModInWrkbkName(ByVal l_comp_name As String)
    LastModInWrkbkName = FSo.GetFileName(Value(VALUE_NAME_LAST_MOD_IN, l_comp_name))
End Function

Friend Property Get LastModOn(Optional ByVal l_comp_name As String = vbNullString) As String
    LastModOn = Value(VALUE_NAME_LAST_MOD_ON, l_comp_name)
End Property

Friend Property Let LastModOn(Optional ByVal l_comp_name As String = vbNullString, _
                                              ByVal l_machine_name As String)
    Value(VALUE_NAME_LAST_MOD_ON, l_comp_name) = l_machine_name
End Property


Private Property Get Value(Optional ByVal v_value_name As String = vbNullString, _
                           Optional ByVal v_v_section_name As String = vbNullString) As String
    Const PROC = "Value/Get"
    
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No value-name provided!"
    If v_v_section_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No component name provided, section name is missing!"
    
    Value = PPFile.Value(v_value_name, v_v_section_name)

End Property

Private Property Let Value(Optional ByVal v_value_name As String, _
                           Optional ByVal v_v_section_name As String = vbNullString, _
                                    ByVal v_value As String)
    Const PROC = "Value/Let"
    
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No value-name provided!"
    If v_v_section_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No component/section name provided, section name is missing!"
    
    PPFile.Value(v_value_name, v_v_section_name) = v_value
                                    
End Property

Public Sub Remove(ByVal r_comp As String)
     PPFile.SectionRemove r_comp
End Sub

Private Function FileHeader() As String
    FileHeader = "Public Common Components. The values are transferred from the Pending.dat file when a pending release" & vbLf & _
                 "Common Component is released to public." & vbLf & _
                 "- LastModAt           : Date/time of the last modification date/time (the export file's creation date repectively)" & vbLf & _
                 "- LastModBy           : User, which had made the last modification" & vbLf & _
                 "- LastModExpFileOrigin: Indicates the 'origin'! of the export file (may point to an export file not available on or not accessable by the used compunter)" & vbLf & _
                 "- LastModIn           : The Workbook/VB-Project in which the last code modification had been made (may point to a Workbook om another computer)" & vbLf & _
                 "- LastModOn           : The computer on which the last modification had been made in the above Workbook."

End Function

Public Function Components() As Dictionary
' ---------------------------------------------------------------------------
' Returns a Dictionary with all sections representing a component.
' ---------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim v   As Variant
    
    With dct
        For Each v In PPFile.SectionNames
            If Not v Like "_*" _
            And Not v Like "@*" _
            Then dct.Add v, vbNullString
        Next v
    End With
    Set Components = dct
    Set dct = Nothing
    
End Function


Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCommonPublic" & "." & sProc
End Function

Public Function Exists(ByVal e_comp_name As String, _
              Optional ByRef e_last_mod_at_datetime_utc As String, _
              Optional ByRef e_last_mod_expfile_fullname_origin As String, _
              Optional ByRef e_last_mod_in_wbk_fullname As String, _
              Optional ByRef e_last_mod_in_wbk_name As String, _
              Optional ByRef e_last_mod_on_machine) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the Private Profile file CommComp.dat has a section named
' (e_comp_name), which indicates an available public Common Component in the
' Common-Components folder.
' ----------------------------------------------------------------------------
    Exists = PPFile.Exists(PPFile.FileName, e_comp_name)
    If Exists Then
        e_last_mod_at_datetime_utc = LastModAt(e_comp_name)
        e_last_mod_expfile_fullname_origin = LastModExpFileOrigin(e_comp_name)
        e_last_mod_in_wbk_fullname = LastModInWrkbkFullName(e_comp_name)
        e_last_mod_in_wbk_name = LastModInWrkbkName(e_comp_name)
        e_last_mod_on_machine = LastModOn(e_comp_name)
    End If
    
End Function

Public Function LastModExpFile(ByVal e_comp As String) As String
' ----------------------------------------------------------------------------
' Provides the full name of a component's (e_comp) last modified Export-File.
' Note: This file's origin is the LastModExpFileOrigin.
' ----------------------------------------------------------------------------
    LastModExpFile = mEnvironment.CommCompsPath & "\" & FSo.GetFileName(LastModExpFileOrigin(e_comp))

End Function

Public Function ExportFiles() As Dictionary
' ----------------------------------------------------------------------------
' Returns all export files in the Common-Components folder as Dictionary with
' the corresponding component name as key and the file's full name as item.
' ----------------------------------------------------------------------------
    
    Dim dct     As New Dictionary
    Dim fld     As Folder
    Dim fle     As File
    Dim sName   As String
    Dim sExt    As String
    
    Set fld = FSo.GetFolder(mEnvironment.CommCompsPath)
    With dct
        For Each fle In fld.Files
            sName = FSo.GetBaseName(fle.Path)
            sExt = FSo.GetExtensionName(fle.Path)
            Select Case sExt
                Case "bas", "cls", "frm"
                    If Not .Exists(sName) _
                    Then .Add sName, fle.Path
            End Select
        Next fle
    End With
    Set ExportFiles = mBasic.KeySort(dct)
    Set dct = Nothing
    
End Function

