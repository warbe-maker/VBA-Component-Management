VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommonServiced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Clss Module clsCommonServiced: Methods and Properties for the maintenance
' ============================== of the Private Profile File CommComps.dat
' in each serviced Workbook's CompMan folder with the following structure:
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
'                              CommComps.dat
' Hosted                     w Provides the hosted components, provided as a
'                              comma delimited string in a Dictionary
'                              (dctHosted) and performs the housekeeping for
'                              the serviced Workbook's CommComps.dat file
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

Public PPFile                                   As New clsPrivProf

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
        
    On Error GoTo eh
    
    With PPFile
        .FileName = mEnvironment.CommCompsServicedPrivProfFileFullName
        .HskpngNames .FileName _
                    , VALUE_NAME_LAST_MOD_AT & "<LastUpdateAtDatTime<LastModAtDatTime<LastModifiedAtDatTime<RawRevisionNumber" _
                    , VALUE_NAME_LAST_MOD_EXP_FILE_ORIG & "<RawExpFileFullName" _
                    , VALUE_NAME_LAST_MOD_IN & "<LastModInWrkbook<LastUpdateModIn<LastUpdateModIn<LastModInWrkbkFullName" _
                    , VALUE_NAME_LAST_MOD_ON & "<LastUpdateModOn<LastModOnMachine" _
                    , "DueModificationWarning>" _
                    , "UsedExportFolder>"
        .FileHeader = FileHeader
        .FileFooter = mEnvironment.PrivateProfileFileFooter
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Class_Terminate()
    Set PPFile = Nothing
End Sub

Private Function FileHeader() As String
    FileHeader = "Common Components used/hosted/private in the serviced Workbook. The values are maintained for any Common Component used or hosted when it is exported due" & vbCrLf & _
                 "to a modification. In case, these values are copied to the pending release registered Common Component." & vbLf & _
                 "- LastModAt           : Date/time of the last modification date/time (the export file's creation date repectively)" & vbCrLf & _
                 "- LastModBy           : User, which had made the last modification" & vbCrLf & _
                 "- LastModExpFileOrigin: Indicates the 'origin'! of the export file (may point to an export file not available on or not accessable by the used compunter)" & vbCrLf & _
                 "- LastModIn           : The Workbook/VB-Project in which the last code modification had been made (may point to a Workbook om another computer)" & vbCrLf & _
                 "- LastModOn           : The computer on which the last modification had been made in the above Workbook."

End Function

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
'    Stop
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
    ErrSrc = "clsCommonServiced" & "." & sProc
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

