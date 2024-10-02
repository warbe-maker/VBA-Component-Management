VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCompManDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Clss Module clsCompManDat: Methods and Properties for the maintenance of
' ========================== the Private Profile File CompMan.dat in each
' serviced Workbook's CompMan folder with the following structure:
'
' [<component-name>]
' KindOfComponent=used|hosted
' LastModAtDateTime=YYYY-MM-DD-hh-mm-ss (UTC)
'
' The entries (sections) are maintained along with the Workbook_BeforeSave
' event via the ExportChangedComponents service. The revision number is
' increased whith each saved modification of a hosted Common Component.
'
' Public Methods:
' ---------------
' ComponentRemove              Removes the section which corresponds with
'                              a Common Component
' IsHostedCommComp
' IsSystemSection
'
' Public Properties:
' ------------------
' Components               r   Dictionary of all components/sections in the
'                              CompMan.dat
' Hosted                     w Provides the hosted components, provided as a
'                              comma delimited string in a Dictionary
'                              (dctHosted) and performs the housekeeping for
'                              the serviced Workbook's CompMan.dat file
'                              file.
' RecentlyUsedExportFolder r/w Keeps a record of the last used export folder.
'                              Used in case the configured name changes.
' KindOfComponent          r/w Information differs between "used", "hosted",
'                              or "private".
' LastModAt     r/w Information saved when a modified Common
'                              Component had been exported.
'
' W. Rauschenberger Berlin, Jul. 2023
' ---------------------------------------------------------------------------
Private Const SECTION_NAME_RECENT_EXPORT        As String = "_MostRecentExport"    ' _ avoids conflict with an existing VBComponent
Private Const VALUE_NAME_REG_STAT_OF_COMPONENT  As String = "KindOfComponent"
Private Const VALUE_NAME_USED_EXPORT_FOLDER     As String = "UsedExportFolder"

Private PPFile      As clsPrivProf

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
        
    On Error GoTo eh
    Dim DatOld As String
    
    mBasic.BoP ErrSrc(PROC)
    If mCompMan.ServicedWrkbk Is Nothing Then Set mCompMan.ServicedWrkbk = ActiveWorkbook
    Set PPFile = New clsPrivProf
    With PPFile
        .FileName = mHskpng.EnvironmentPrivProfFileFullName
    
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
                   , VALUE_NAME_LAST_MOD_AT & "<LastUpdateAtDatTime" _
                   , VALUE_NAME_LAST_MOD_IN & "<LastModInWrkbook<LastUpdateModIn" _
                   , VALUE_NAME_LAST_MOD_ON & "<LastUpdateModOn" _
                   , "DueModificationWarning>" _
                   , "LastUpdateAtDatTime>" _
                   , "LastUpdateModIn>" _
                   , "LastUpdateModOn>" _
                   , "LastModifiedAtDatTime>" & VALUE_NAME_LAST_MOD_AT _
                   , "LastModAtDatTime>LastModAt" _
                   , "LastModOnMachine>LastModOn" _
                   , "LastModInWrkbkFullName>LastModIn"
                   
        If .FileHeader <> FileHeader Then
            .FileHeader = FileHeader
        End If
    End With
    
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Class_Terminate()
    Set PPFile = Nothing
End Sub

Public Property Let CompName(ByVal c_comp As String)
    PPFile.Section = c_comp
End Property

Friend Property Let KindOfComponent(Optional ByVal comp_name As String = vbNullString, _
                                               ByVal comp_reg_state As enKindOfComp)
    Value(comp_name, VALUE_NAME_REG_STAT_OF_COMPONENT) = mCompMan.RegState(comp_reg_state)
End Property

Friend Property Get KindOfComponent(Optional ByVal comp_name As String) As enKindOfComp
    KindOfComponent = mCompMan.RegState(Value(comp_name, VALUE_NAME_REG_STAT_OF_COMPONENT))
End Property

Friend Property Get LastModAt(Optional ByVal l_comp As String) As String
    LastModAt = Value(l_comp, VALUE_NAME_LAST_MOD_AT)
End Property

Friend Property Let LastModAt(Optional ByVal l_comp As String, _
                                                  ByVal l_datetime_utc As String)
    Value(l_comp, VALUE_NAME_LAST_MOD_AT) = l_datetime_utc
End Property

Friend Property Get LastModBy(Optional ByVal l_comp As String) As String
    LastModBy = Value(l_comp, VALUE_NAME_LAST_MOD_BY)
End Property

Friend Property Let LastModBy(Optional ByVal l_comp As String, _
                                                  ByVal l_datetime_utc As String)
    Value(l_comp, VALUE_NAME_LAST_MOD_BY) = l_datetime_utc
End Property

Friend Property Get LastModExpFileOrigin(Optional ByVal m_comp As String = vbNullString) As String
    LastModExpFileOrigin = Value(m_comp, VALUE_NAME_LAST_MOD_EXP_FILE_ORIG)
End Property

Friend Property Let LastModExpFileOrigin(Optional ByVal m_comp As String = vbNullString, _
                                            ByVal p_mod_exp_file As String)
    Value(m_comp, VALUE_NAME_LAST_MOD_EXP_FILE_ORIG) = p_mod_exp_file
End Property

Friend Property Get LastModInWrkbkFullName(Optional ByVal l_comp As String) As String
    LastModInWrkbkFullName = Value(l_comp, VALUE_NAME_LAST_MOD_IN)
End Property

Friend Property Let LastModInWrkbkFullName(Optional ByVal l_comp As String, _
                                            ByVal l_datetime_utc As String)
    Value(l_comp, VALUE_NAME_LAST_MOD_IN) = l_datetime_utc
End Property

Friend Property Get LastModOn(Optional ByVal l_comp As String) As String
    LastModOn = Value(l_comp, VALUE_NAME_LAST_MOD_ON)
End Property

Friend Property Let LastModOn(Optional ByVal l_comp As String, _
                                              ByVal l_datetime_utc As String)
    Value(l_comp, VALUE_NAME_LAST_MOD_ON) = l_datetime_utc
End Property

Public Property Get NameHskpngSection() As String:         NameHskpngSection = PPFile.NameHskpngSection:    End Property

Friend Property Get RecentlyUsedExportFolder() As String
    RecentlyUsedExportFolder = Value(SECTION_NAME_RECENT_EXPORT, VALUE_NAME_USED_EXPORT_FOLDER)
End Property

Friend Property Let RecentlyUsedExportFolder(ByVal s As String)
    Value(SECTION_NAME_RECENT_EXPORT, VALUE_NAME_USED_EXPORT_FOLDER) = s
End Property

Private Property Get Value(Optional ByVal v_section_name As String = vbNullString, _
                           Optional ByVal v_value_name As String = vbNullString) As String
    Const PROC = "Value/Get"
    
    If v_section_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No component-name provided!"
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    
    Value = PPFile.Value(v_value_name, v_section_name)

End Property

Private Property Let Value(Optional ByVal v_section_name As String = vbNullString, _
                           Optional ByVal v_value_name As String = vbNullString, _
                                    ByVal v_value As String)
    Const PROC = "Value/Let"
    
    If v_section_name = vbNullString _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No component-name provided!"
    If v_value_name = vbNullString _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No value-name provided!"
    
    PPFile.Value(v_value_name, v_section_name) = v_value
                                    
End Property

Public Sub ComponentRemove(ByVal r_comp_name As String)
    PPFile.SectionRemove r_comp_name
End Sub

Public Function Components() As Dictionary
' ---------------------------------------------------------------------------
' Returns a Dictionary with all sections representing a component on the
' serviced Workbook (internal sections are excluded).
' ---------------------------------------------------------------------------
    Dim dct As Dictionary
    Dim v   As Variant
    
    Set dct = PPFile.SectionNames
    For Each v In dct
        If v Like "_*" Then dct.Remove v
        If v Like "@*" Then dct.Remove v
    Next v
    Set Components = dct
    Set dct = Nothing
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCompManDat" & "." & sProc
End Function

Private Function FileHeader() As String
    FileHeader = "Used or hosted Common Components in the corresponding serviced Workbook """ & Serviced.Wrkbk.Name & """." & vbLf & _
                 "The values are provided when a Common Component's code has been modified in the serviced Workbook or" & vbLf & _
                 "when a used or hosted Common Component has been updated with code modified elsewhere." & vbLf & _
                 "- LastModAt           : Date/time of the last modification date/time (the export file's creation date repectively)" & vbLf & _
                 "- LastModBy           : User, which had made the last modification" & vbLf & _
                 "- LastModExpFileOrigin: Indicates the 'origin'! of the export file (may point to an export file not available on or not accessable by the used compunter)" & vbLf & _
                 "- LastModIn           : The Workbook/VB-Project in which the last code modification had been made (may point to a Workbook om another computer)" & vbLf & _
                 "- LastModOn           : The computer on which the last modification had been made in the above Workbook."

End Function

Public Function IsHostedCommComp(ByVal i_comp As String, _
                        Optional ByRef i_last_modified As String) As Boolean
    
    If Components.Exists(i_comp) Then
        IsHostedCommComp = True
        i_last_modified = LastModAt(i_comp)
    End If
    
End Function

Public Function IsSystemSection(ByVal i_sect As String) As Boolean
' ---------------------------------------------------------------------------
' Returns TRUE when the section (i_sect) is an internal system/administration
' section.
' ---------------------------------------------------------------------------

    IsSystemSection = i_sect = SECTION_NAME_RECENT_EXPORT _
    Or PPFile.IsSystemSection(i_sect)
    
End Function

Public Function IsUsedCommComp(ByVal i_comp As String, _
                      Optional ByRef i_last_modified As String) As Boolean
    
    If Components.Exists(i_comp) Then
        IsUsedCommComp = True
        i_last_modified = LastModAt(i_comp)
    End If
    
End Function

Public Function LastModInWrkbkName(ByVal l_comp As String) As String
    LastModInWrkbkName = FSo.GetFileName(Value(l_comp, VALUE_NAME_LAST_MOD_IN))
End Function

Public Sub RemoveComponent(ByVal r_comp As String)
    PPFile.SectionRemove r_comp
End Sub

