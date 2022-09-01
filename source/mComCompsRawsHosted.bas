Attribute VB_Name = "mComCompsRawsHosted"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mComCompsRawsHosted
' Maintains in a file named ComCompsHosted.dat for each Workbook which hosts
' at least one Raw Common Component with the following structure:
'
' [<component-name]
' ExpFileFullName=<export-file-full-name>
' RevisionNumber=yyyy-mm-dd.nnn
'
' The entries (sections) are maintained along with the Workbook_BeforeSave
' event via the ExportChangedComponents service. The revision number is
' increased whith each saved modification.
' ---------------------------------------------------------------------------
Private Const VNAME_RAW_REVISION_NUMBER     As String = "RawRevisionNumber"
Private Const VNAME_RAW_EXP_FILE_FULL_NAME  As String = "RawExpFileFullName"

Public Function IsNewHostedRaw(ByVal in_vbc_name As String) As Boolean
' ---------------------------------------------------------------------------
' Returns TRUE when the serviced Workbook claims a component "hosted raw
' common coponent" and the concerned component is yet not registered in the
' Serviced Workbook's 'ComCompsHosted.dat' (ComCompsHostedFileFullName) or no
' 'ComCompsHosted.dat' exists.
' ---------------------------------------------------------------------------
    IsNewHostedRaw = Not IsRegisteredLocally(in_vbc_name)
End Function

Public Property Get IsRegisteredLocally(ByVal irl_vbc_name As String) As Boolean
    IsRegisteredLocally = Exists(irl_vbc_name)
End Property

Public Sub SaveToGlobalFolder(ByVal stgf_vbc_name As String, _
                              ByVal stgf_exp_file As File, _
                              ByVal stgf_exp_file_full_name As String)
' ------------------------------------------------------------------------------
' Save a copy of the hosted raw`s (stgf_vbc_name) export file to the Common
' Components folder which serves as source for the update of Common Components
' used in other VB-Projects.
' ------------------------------------------------------------------------------
    Dim frxFile As File
    Dim fso     As New FileSystemObject
    
    mComCompsRawsGlobal.SavedExpFile(stgf_vbc_name) = stgf_exp_file
    '~~ When the Export file has a .frm extension the .frx file needs to be copied too
    If fso.GetExtensionName(stgf_exp_file_full_name) = "frm" Then
        Set frxFile = fso.GetFile(Replace(stgf_exp_file_full_name, "frm", "frx"))
        mComCompsRawsGlobal.SavedExpFile(stgf_vbc_name) = frxFile
    End If

    mComCompsRawsGlobal.RawExpFileFullName(stgf_vbc_name) = stgf_exp_file_full_name
    mComCompsRawsGlobal.RawHostWbBaseName(stgf_vbc_name) = fso.GetBaseName(mService.Serviced.FullName)
    mComCompsRawsGlobal.RawHostWbFullName(stgf_vbc_name) = mService.Serviced.FullName
    mComCompsRawsGlobal.RawHostWbName(stgf_vbc_name) = mService.Serviced.Name
    mComCompsRawsGlobal.RawSavedRevisionNumber(stgf_vbc_name) = mComCompsRawsHosted.RawRevisionNumber(stgf_vbc_name)
    
    Set fso = Nothing
End Sub

Public Property Get IsRegisteredGlobally(ByVal irg_vbc_name As String) As Boolean
    IsRegisteredGlobally = mComCompsRawsGlobal.Exists(irg_vbc_name)
End Property

Public Property Get ComCompsHostedFileFullName() As String
    Dim wbk As Workbook
    Dim fso As New FileSystemObject
    
    Set wbk = mService.Serviced
    ComCompsHostedFileFullName = Replace(wbk.FullName, wbk.Name, "ComCompsHosted.dat")
    If Not fso.FileExists(ComCompsHostedFileFullName) Then
        fso.CreateTextFile ComCompsHostedFileFullName
    End If
    Set fso = Nothing
    
End Property

Public Property Get RawExpFileFullName(Optional ByVal comp_name As String) As String
    RawExpFileFullName = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME)
End Property

Public Property Let RawExpFileFullName(Optional ByVal comp_name As String, _
                                                ByVal exp_file_full_name As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_EXP_FILE_FULL_NAME) = exp_file_full_name
End Property

Public Property Get RawRevisionNumber(Optional ByVal comp_name As String) As String
    RawRevisionNumber = Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER)
End Property

Public Property Let RawRevisionNumber(Optional ByVal comp_name As String, _
                                               ByVal comp_rev_no As String)
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER) = comp_rev_no
End Property

Private Property Get Value(Optional ByVal pp_section As String, _
                           Optional ByVal pp_value_name As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' file ComCompsHostedFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    Value = mFile.Value(pp_file:=ComCompsHostedFileFullName _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Let Value(Optional ByVal pp_section As String, _
                           Optional ByVal pp_value_name As String, _
                                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (pp_value) under the name (pp_value_name) into the
' ComCompsHostedFileFullName.
' ----------------------------------------------------------------------------
    Const PROC = "Value_Let"
    
    On Error GoTo eh
    mFile.Value(pp_file:=ComCompsHostedFileFullName _
              , pp_section:=pp_section _
              , pp_value_name:=pp_value_name _
               ) = pp_value

xt: Exit Property

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Function Components() As Dictionary
    Set Components = mFile.SectionNames(ComCompsHostedFileFullName)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mRaw" & "." & sProc
End Function

Public Function Exists(ByVal raw_vbc_name As String) As Boolean
    Exists = Components.Exists(raw_vbc_name)
End Function

Public Function MaxRawLenght() As Long
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

Public Sub Remove(ByVal comp_name As String)
    mFile.RemoveSections pp_file:=ComCompsHostedFileFullName _
                       , pp_sections:=comp_name
End Sub

Public Sub RawRevisionNumberIncrease(ByVal comp_name As String)
' ----------------------------------------------------------------------------
' Increases the revision number by one starting with 1 for a new day.
' ----------------------------------------------------------------------------
    Dim RevNo   As Long
    Dim RevDate As String
    
    If RawRevisionNumber(comp_name) = vbNullString Then
        RevNo = 1
    Else
        RevNo = Split(RawRevisionNumber(comp_name), ".")(1)
        RevDate = Split(RawRevisionNumber(comp_name), ".")(0)
        If RevDate <> Format(Now(), "YYYY-MM-DD") _
        Then RevNo = 1 _
        Else RevNo = RevNo + 1
    End If
    Value(pp_section:=comp_name, pp_value_name:=VNAME_RAW_REVISION_NUMBER) = Format(Now(), "YYYY-MM-DD") & "." & Format(RevNo, "000")

End Sub

