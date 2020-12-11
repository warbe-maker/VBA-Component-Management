Attribute VB_Name = "mRaws"
Option Explicit

Private Property Get FileName() As String
    FileName = mCfg.CompManAddinPath & "\Raws.dat"
End Property

Public Property Get RawExpFile( _
         Optional ByVal raw_wbk_name As String, _
         Optional ByVal raw_comp_name As String) As String
    RawExpFile = mFl.Value(vl_file:=FileName, vl_section:=raw_wbk_name, vl_value_name:=raw_comp_name)
End Property

Public Property Let RawExpFile( _
         Optional ByVal raw_wbk_name As String, _
         Optional ByVal raw_comp_name As String, _
                  ByVal raw_exp_file_full_name As String)
    mFl.Value(vl_file:=FileName, vl_section:=raw_wbk_name, vl_value_name:=raw_comp_name) = raw_exp_file_full_name
End Property

Public Property Get IsRaw( _
           Optional ByVal raw_comp_name As String) As Boolean
    IsRaw = mFl.Values(vl_file:=FileName).Exists(raw_comp_name)
End Property

