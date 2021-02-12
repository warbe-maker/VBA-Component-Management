Attribute VB_Name = "mSync"
Option Explicit

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSync." & s
End Function

Public Sub VbProject( _
               ByVal clone_project As Workbook, _
      Optional ByVal raw_project As String = vbNullString)
' --------------------------------------------------------
'
' --------------------------------------------------------
    Const PROC = "VbProject"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim cClone  As clsComp
    Dim cRaw    As clsRaw
    Dim vbc     As VBComponent
    
    '~~ Update all clone componentes where the raw had changed
    '~~ and remove all components where there is no Export File
    For Each vbc In clone_project.VbProject.VBComponents
        Set cClone = New clsComp
        Set cRaw = New clsRaw
        With cClone
            .Wrkbk = clone_project
            .CompName = vbc.name
            .VBComp = vbc
            cRaw.HostFullName = raw_project
            cRaw.CompName = .CompName
            cRaw.ExpFileExtension = .ExpFileExtension
            cRaw.CloneExpFileFullName = .ExpFilePath
            If Not fso.FileExists(cRaw.ExpFilePath) Then
                '~~ The corresponding raw's Export File does not exist which indicates that the component no longer exists
                '~~ This is cruical in case the raw component just never had been exported! However, since the registration
                '~~ of the raw VB-Project is done by a service which makes sure all comopnents are exorted this should never
                '~~ be the case
                .Wrkbk.VbProject.VBComponents.Remove .VBComp
                GoTo next_vbc
            End If
            If .Changed Then .VBComp.Export .ExpFilePath ' make sure the current code version is exported
            If cRaw.Changed Then
                mRenew.ByImport rn_wb:=.Wrkbk _
                              , rn_comp_name:=.CompName _
                              , rn_exp_file_full_name:=cRaw.ExpFilePath
            End If
        End With
        
        Set cClone = Nothing
        Set cRaw = Nothing
next_vbc:
    Next vbc
    
xt: Set fso = Nothing
    Exit Sub

eh:
End Sub
