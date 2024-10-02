VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPPCompManDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Clss Module clsCompManDat: Methods and Properties for the maintenance of
' ========================== the Private Profile File CompMan.dat in each
' serviced Workbook's parent folder with the following structure:
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
' RecentlyUsedExportFolder r/w .
' RegistrationState        r/w .
' LastModAtDateTimeUTC           r/w .
'
' W. Rauschenberger Berlin, Jul. 2023
' ---------------------------------------------------------------------------
Private Const SECTION_NAME_RECENT_EXPORT        As String = "_MostRecentExport" ' _ avoids conflict with an existing VBComponent
Private Const VALUE_NAME_LAST_MOD_DATE_TIME     As String = "LastModifiedAtDatTime"
Private Const VALUE_NAME_REG_STAT_OF_COMPONENT  As String = "KindOfComponent"
Private Const VALUE_NAME_USED_EXPORT_FOLDER     As String = "UsedExportFolder"

Private dctHosted                               As Dictionary
Public PP                                       As clsPrivProf

Private Sub Class_Initialize()
    Const PROC = "Class_Initialize"
        
    Dim DatOld As String
    Dim DatNew As String
    
    mBasic.BoP ErrSrc(PROC)
    Set PP = New clsPrivProf
    DatOld = Serviced.Wrkbk.Path & "\CompMan.dat"
    DatNew = Serviced.Folder & "\CompMan.dat" ' this will create the folder when yet not existing
    If FSo.FileExists(DatOld) Then
        FSo.MoveFile DatOld, DatNew
    End If
    
    With PP
        .FileName = DatNew ' the default throughout all methods and properties
    
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
                   , VALUE_NAME_LAST_MOD_DATE_TIME & "<LastModAtDateTime<RevisionNumber<RawRevisionNumber" _
                   , "DueModificationWarning>"
    End With
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Private Sub Class_Terminate()
    Set PP = Nothing
End Sub

Friend Property Get Components() As Dictionary: Set Components = PP.SectionNames: End Property

Private Property Get CompRegState(Optional ByVal comp_name As String) As enCommCompRegState
    CompRegState = CompMan.CommCompRegStateEnum(PP.Value(name_section:=comp_name, name_value:=VALUE_NAME_REG_STAT_OF_COMPONENT))
End Property

Private Property Let CompRegState(Optional ByVal comp_name As String, _
                                           ByVal comp_reg_state As enCommCompRegState)
    PP.Value(name_section:=comp_name, name_value:=VALUE_NAME_REG_STAT_OF_COMPONENT) = CompMan.CommCompRegStateString(comp_reg_state)
End Property

Friend Property Let Hosted(ByVal s As String)
' ----------------------------------------------------------------------------
' Provides the hosted components, provided as a comma delimited string in a
' Dictionary (dctHosted) and performs the housekeeping for Common Components
' in the serviced Workbook's CompMan.dat file.
' ----------------------------------------------------------------------------
    Dim v       As Variant
    Dim sComp   As String
    Dim dct     As New Dictionary
    
    Set dctHosted = Nothing: Set dctHosted = New Dictionary
    For Each v In Split(s, ",")
        sComp = Trim$(v)
        If Not dct.Exists(sComp) Then
            dctHosted.Add sComp, sComp
        End If
    Next v
    
End Property

Friend Property Get LastModAtDateTimeUTC(Optional ByVal r_comp_name As String) As String
    LastModAtDateTimeUTC = PP.Value(name_section:=r_comp_name, name_value:=VALUE_NAME_LAST_MOD_DATE_TIME)
End Property

Friend Property Let LastModAtDateTimeUTC(Optional ByVal r_comp_name As String, _
                                                  ByVal r_rev_no As String)
    PP.Value(name_section:=r_comp_name, name_value:=VALUE_NAME_LAST_MOD_DATE_TIME) = r_rev_no
End Property

Friend Property Get RecentlyUsedExportFolder() As String
    RecentlyUsedExportFolder = PP.Value(name_section:=SECTION_NAME_RECENT_EXPORT, name_value:=VALUE_NAME_USED_EXPORT_FOLDER)
End Property

Friend Property Let RecentlyUsedExportFolder(ByVal s As String)
    PP.Value(name_section:=SECTION_NAME_RECENT_EXPORT, name_value:=VALUE_NAME_USED_EXPORT_FOLDER) = s
End Property

Friend Property Get RegistrationState(Optional ByVal comp_name As String) As enCommCompRegState
    RegistrationState = CompRegState(comp_name)
End Property

Friend Property Let RegistrationState(Optional ByVal comp_name As String, _
                                               ByVal comp_reg_state As enCommCompRegState)
    CompRegState(comp_name) = comp_reg_state
End Property

Friend Property Get RegStateString(Optional ByVal comp_name As String) As String
    Select Case RegistrationState(comp_name)
        Case enRegStateHosted:  RegStateString = "hosted"
        Case enRegStatePrivate: RegStateString = "private"
        Case enRegStateUsed:    RegStateString = "used"
    End Select
End Property

Public Sub ComponentRemove(ByVal r_comp_name As String)
    PP.SectionRemove r_comp_name
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCompManDat" & "." & sProc
End Function

Public Sub Hskpng()
' ------------------------------------------------------------------------------
' Removes obsolete sections which are those neither representing an existing
' VBComponent no another valid section's Name.
' ------------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    HskpngObsolete
    HskpngHosted

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngHosted()
' ------------------------------------------------------------------------------
' Maintains in the serviced Workbook's CompMan.dat file for any component
' claimed hosted the corresponding RegistrationState and the Revision-Number
' (when missing).
' ------------------------------------------------------------------------------
    Const PROC      As String = "HskpngHosted"
    
    On Error GoTo eh
    Dim wbk         As Workbook
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.ServicedWbk
    
    For Each v In dctHosted
        If mComp.Exists(v, wbk) Then
            If RegistrationState(v) <> enRegStateHosted Then
                RegistrationState(v) = enRegStateHosted
            End If
            If LastModAtDateTimeUTC(v) = vbNullString Then
                LastModAtDateTimeUTC(v) = CompMan.UTC & " (UTC)"
            End If
        End If
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngNotHosted()
' ----------------------------------------------------------------------------
' When a former hosting Workbook not or no longer claims a Common Component
' hosted the RegistrationState in the serviced Workbook's CompMan.dat file
' is changed to enRegStateUsed.
' ----------------------------------------------------------------------------
    Const PROC      As String = "HskpngNotHosted"
    
    On Error GoTo eh
    Dim vCompName           As Variant
    Dim wbk         As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.ServicedWbk
    For Each vCompName In Components
        If Not dctHosted.Exists(vCompName) Then
            If mComp.Exists(vCompName, wbk) Then
                RegistrationState(vCompName) = enRegStateUsed
            End If
        End If
    Next vCompName

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub HskpngObsolete()
' ------------------------------------------------------------------------------
' Remove sections representing VBComponents no longer existing and those with an
' invalid name.
' ------------------------------------------------------------------------------
    Const PROC = "HskpngObsolete"
    
    On Error GoTo eh
    Dim v   As Variant
    Dim wbk As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.ServicedWbk
    For Each v In Components
        If HskpngSectionIsInvalid(v, wbk) Then
            ComponentRemove v
        End If
    Next v
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function HskpngSectionIsInvalid(ByVal h_section As String, _
                                        ByVal h_wbk As Workbook) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the section (h_section) is neither an existing
' VBComponent's Name nor another valid section's Name.
' ------------------------------------------------------------------------------
    Select Case True
        Case mComp.Exists(h_section, h_wbk)
        Case h_section = SECTION_NAME_RECENT_EXPORT
        Case h_section = PP.NameHskpngSection
        Case Else
            HskpngSectionIsInvalid = True
    End Select
    
End Function

Private Function MaxRawLenght() As Long
' -----------------------------------------------
' Returns the max length of a raw componen's name
' -----------------------------------------------
    Const PROC = "MaxRawLenght"
    
    On Error GoTo eh
    Dim v As Variant
    
    For Each v In Components
        MaxRawLenght = Max(MaxRawLenght, Len(v))
    Next v
    
xt: Exit Function

eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

