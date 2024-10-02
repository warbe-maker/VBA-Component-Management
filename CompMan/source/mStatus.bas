Attribute VB_Name = "mStatus"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mStatus: Provides various status results.
' ------------------------
'
' ----------------------------------------------------------------------------

Public Function IsPendingReleaseCommComp(ByVal i_comp As String, _
                                Optional ByRef e_mod_at_datetime As String, _
                                Optional ByRef e_mod_export_file_name As String, _
                                Optional ByRef e_mod_in_wbk_full_name As String, _
                                Optional ByRef e_mod_in_wbk_name As String, _
                                Optional ByRef e_mod_on_machine As String) As Boolean
    
    If CommonPublic Is Nothing Then Set CommonPublic = New clsCommonPublic
    If CommonPending Is Nothing Then Set CommonPending = New clsCommonPending
    IsPendingReleaseCommComp = CommonPending.Exists(i_comp _
                                                  , e_mod_at_datetime _
                                                  , e_mod_export_file_name _
                                                  , e_mod_in_wbk_full_name _
                                                  , e_mod_in_wbk_name _
                                                  , e_mod_on_machine)
    
End Function

Public Function IsUsedCommonComp(ByVal i_comp As String, _
                        Optional ByRef i_last_mod_at_datetime As String) As Boolean
    
    If Serviced Is Nothing Then Set Serviced = New clsServiced
    Serviced.Wrkbk = ThisWorkbook
    If PPCompManDat Is Nothing Then Set PPCompManDat = New clsCompManDat
    IsUsedCommonComp = PPCompManDat.IsUsedCommComp(i_comp, i_last_mod_at_datetime) = True
    
End Function

Private Sub Test_IsUsedCommonComp()

    Dim sLastModAtDatetime As String
    If IsUsedCommonComp("mBasic", sLastModAtDatetime) Then
        Debug.Print "Is a used Common Component"
        Debug.Print "Last modified at: " & sLastModAtDatetime
    End If
    
End Sub

Private Sub Test_IsPendingReleaseCommComp()
    Dim sModAtDatetime      As String
    Dim sModExportFileName  As String
    Dim sModInWbkFullName   As String
    Dim sModInWbkName       As String
    Dim sModOnMachine       As String
    Dim b                   As Boolean
    
    If IsPendingReleaseCommComp("mBasic" _
                              , sModAtDatetime _
                              , sModExportFileName _
                              , sModInWbkFullName _
                              , sModInWbkName _
                              , sModOnMachine) Then
        Debug.Print "Is not pending release"
        Debug.Print "ModAtDatetime ...: " & sModAtDatetime
        Debug.Print "ModExportFileName: " & sModExportFileName
        Debug.Print "ModInWbkFullName : " & sModInWbkFullName
        Debug.Print "ModInWbkame .....: " & sModInWbkName
        Debug.Print "ModOnMachine ....: " & sModOnMachine
    Else
        Debug.Print "Is not pending release"
    End If
End Sub

Public Function IsPublicCommonComponent(ByVal i_comp As String, _
                               Optional ByRef i_export_file_extention As String, _
                               Optional ByRef i_last_mod_atdatetime_utc As String, _
                               Optional ByRef i_last_mod_expfile_fullname_origin As String, _
                               Optional ByRef i_last_mod_inwbk_fullname As String, _
                               Optional ByRef i_last_mod_inwbk_name As String, _
                               Optional ByRef i_last_mod_on_machine As String) As Boolean
' ----------------------------------------------------------------------------
' When the component (i_comp) exists in the CommComps.dat Private Profile File
' the function returns TRUE inxluding all relevant values.
' ----------------------------------------------------------------------------

    If CommonPublic Is Nothing Then Set CommonPublic = New clsCommonPublic
    IsPublicCommonComponent = CommonPublic.Exists(i_comp _
                                                , i_export_file_extention _
                                                , i_last_mod_atdatetime_utc _
                                                , i_last_mod_expfile_fullname_origin _
                                                , i_last_mod_inwbk_fullname _
                                                , i_last_mod_inwbk_name _
                                                , i_last_mod_on_machine)
                        
End Function

