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
' serviced Workbook's parent folder with the following structure:
'
' [<component-name>]
' KindOfComponent=used|hosted
' RevisionNumber=yyyy-mm-dd.nnn
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
' RevisionNumber           r/w .
'
' W. Rauschenberger Berlin, Jul. 2023
' ---------------------------------------------------------------------------

Private Const SECTION_NAME_RECENT_EXPORT            As String = "_MostRecentExport" ' _ avoids conflict with an existing VBComponent
Private Const VALUE_NAME_LAST_MOD_REVISION_NUMBER   As String = "VALUE_NAME_LAST_MOD_REVISION_NUMBER"
Private Const VALUE_NAME_REG_STAT_OF_COMPONENT      As String = "VALUE_NAME_REG_STAT_OF_COMPONENT"
Private Const VALUE_NAME_USED_EXPORT_FOLDER         As String = "VALUE_NAME_USED_EXPORT_FOLDER"
Private PP                                          As clsPP
Private fso                                         As FileSystemObject
Private dctHosted                                   As Dictionary

Private Sub Class_Initialize()
    Dim wbk     As Workbook
    Dim dct     As New Dictionary
    
    Set PP = New clsPP
    
    '~~> Private Property File Full Name
    Set wbk = Services.Serviced
    PP.FullName = Replace(wbk.FullName, wbk.Name, "CompMan.dat")
    
    '~~> Value Names with the syntax (allows the on-the-fly-change of value names as well as their removal):
    '~~> "<current>"                 = no action
    '~~> "<current-name>:<old-name>" = rename in all sections
    '~~> ":<remove-name>"            = remove in all sections
    dct.Add VALUE_NAME_LAST_MOD_REVISION_NUMBER, "RevisionNumber:RawRevisionNumber"
    dct.Add VALUE_NAME_REG_STAT_OF_COMPONENT, "KindOfComponent"
    dct.Add VALUE_NAME_USED_EXPORT_FOLDER, "UsedExportFolder"
    dct.Add "Obsolete_1", ":DueModificationWarning"
    PP.Names = dct
    Set dct = Nothing
    
End Sub

Private Sub Class_Terminate()
    Set fso = Nothing
    Set PP = Nothing
End Sub

Friend Property Get Components() As Dictionary
    Set Components = PP.Components
End Property

Private Property Get CompRegState(Optional ByVal comp_name As String) As enCommCompRegState
    With PP
        CompRegState = CompMan.CommCompRegStateEnum(.Value(pp_section:=comp_name, pp_value_name:=.Name(VALUE_NAME_REG_STAT_OF_COMPONENT)))
    End With
End Property

Private Property Let CompRegState(Optional ByVal comp_name As String, _
                                           ByVal comp_reg_state As enCommCompRegState)
    With PP
        .Value(pp_section:=comp_name, pp_value_name:=.Name(VALUE_NAME_REG_STAT_OF_COMPONENT)) = CompMan.CommCompRegStateString(comp_reg_state)
    End With
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
    Hskpng
    
End Property

Friend Property Get RecentlyUsedExportFolder() As String
    With PP
        RecentlyUsedExportFolder = .Value(pp_section:=SECTION_NAME_RECENT_EXPORT, pp_value_name:=.Name(VALUE_NAME_USED_EXPORT_FOLDER))
    End With
End Property

Friend Property Let RecentlyUsedExportFolder(ByVal s As String)
    With PP
        .Value(pp_section:=SECTION_NAME_RECENT_EXPORT, pp_value_name:=.Name(VALUE_NAME_USED_EXPORT_FOLDER)) = s
    End With
End Property

Friend Property Get RegistrationState(Optional ByVal comp_name As String) As enCommCompRegState
    RegistrationState = CompRegState(comp_name)
End Property

Friend Property Let RegistrationState(Optional ByVal comp_name As String, _
                                               ByVal comp_reg_state As enCommCompRegState)
    CompRegState(comp_name) = comp_reg_state
End Property

Friend Property Get RevisionNumber(Optional ByVal r_comp_name As String) As String
' ----------------------------------------------------------------------------
' Returns the revision number in the format YYYY-MM-DD.n
' ----------------------------------------------------------------------------
    With PP
        RevisionNumber = .Value(pp_section:=r_comp_name, pp_value_name:=.Name(VALUE_NAME_LAST_MOD_REVISION_NUMBER))
    End With
End Property

Friend Property Let RevisionNumber(Optional ByVal r_comp_name As String, _
                                            ByVal r_rev_no As String)
    With PP
        .Value(pp_section:=r_comp_name, pp_value_name:=.Name(VALUE_NAME_LAST_MOD_REVISION_NUMBER)) = r_rev_no
    End With
End Property

Public Function CommCompUsedIsKnown(ByVal comp_name As String) As Boolean
    CommCompUsedIsKnown = ComponentsRegistered(enRegStateUsed).Exists(comp_name)
End Function

Public Sub ComponentRemove(ByVal r_comp_name As String)
    PP.ComponentRemove r_comp_name
End Sub

Private Function ComponentsRegistered(ByVal c_reg_state As enCommCompRegState) As Dictionary
    Dim vSection    As Variant
    Dim dct         As New Dictionary
    
    With PP
        For Each vSection In PP.Components
            If .Value(pp_section:=vSection, pp_value_name:=.Name(VALUE_NAME_REG_STAT_OF_COMPONENT)) = CompMan.CommCompRegStateString(c_reg_state) _
            Then dct.Add vSection, vbNullString
        Next vSection
    End With
    Set ComponentsRegistered = dct
    Set dct = Nothing
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsCompManDat" & "." & sProc
End Function

Private Sub Hskpng()
' ------------------------------------------------------------------------------
' Removes obsolete sections which are those neither representing an existing
' VBComponent no another valid section's Name.
' ------------------------------------------------------------------------------
    Const PROC = "Hskpng"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    HskpngObsolete
    HskpngHosted
    PP.Reorg

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    Dim Comp        As clsComp
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    
    For Each v In dctHosted
        If mComp.Exists(v, wbk) Then
            If Me.RegistrationState(v) <> enRegStateHosted _
            Then Me.RegistrationState(v) = enRegStateHosted
            If Me.RevisionNumber(v) = vbNullString _
            Then Me.RevisionNumber = Me.RevisionNumberInitial
        End If
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    Dim v           As Variant
    Dim wbk         As Workbook
    
    mBasic.BoP ErrSrc(PROC)
    Set wbk = Services.Serviced
    For Each v In Components
        If Not dctHosted.Exists(v) Then
            If mComp.Exists(v, wbk) Then
                CompManDat.RegistrationState(v) = enRegStateUsed
            End If
        End If
    Next v

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
    Set wbk = Services.Serviced
    For Each v In PP.Components
        If HskpngSectionIsInvalid(v, wbk) Then
            PP.ComponentRemove v
        End If
    Next v
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function HskpngSectionIsInvalid(ByVal h_section As String, _
                                        ByVal h_wbk As Workbook) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the section (h_section) is invalid, which is the case when
' it is neither an existing VBComponent's Name nor another valid section's Name.
' ------------------------------------------------------------------------------
    Select Case True
        Case mComp.Exists(h_section, h_wbk)
        Case h_section = SECTION_NAME_RECENT_EXPORT
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Reorg()
    PP.Reorg
End Sub

Public Function RevisionNumberInitial() As String
' ----------------------------------------------------------------------------
' Returns an initial revision number in the form: YYYY-MM-DD.001
' ----------------------------------------------------------------------------
    RevisionNumberInitial = Format(Now(), "YYYY-MM-DD") & ".001"
End Function

