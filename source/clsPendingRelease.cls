VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPendingRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ---------------------------------------------------------------------------
' Class Module clsRelease: Properties for the registration/maintenance of
' ======================== pending Releases.
'
' ---------------------------------------------------------------------------
Friend Property Get ModifiedAtDateTime(Optional ByVal m_name As String) As String
    ModifiedAtDateTime = CommComps.PendingReleaseModAtDateTime(m_name)
End Property

Friend Property Let ModifiedAtDateTime(Optional ByVal m_name As String, _
                                                  ByVal m_mod_last As String)
    CommComps.PendingReleaseModAtDateTime(m_name) = m_mod_last
End Property

Friend Property Get ModifiedInWbkFullName(Optional ByVal m_name As String) As String
    ModifiedInWbkFullName = CommComps.PendingReleaseModInWbkFullName(m_name)
End Property

Friend Property Let ModifiedInWbkFullName(Optional ByVal m_name As String, _
                                               ByVal m_mod_by_wbk_name As String)
    CommComps.PendingReleaseModInWbkFullName(m_name) = m_mod_by_wbk_name
End Property

Friend Property Get ModifiedInWbkName(Optional ByVal m_name As String) As String
    ModifiedInWbkName = CommComps.PendingReleaseModInWbkName(m_name)
End Property

Friend Property Let ModifiedInWbkName(Optional ByVal m_name As String, _
                                               ByVal m_mod_by_wbk_name As String)
    CommComps.PendingReleaseModInWbkName(m_name) = m_mod_by_wbk_name
End Property

Friend Property Get ModifiedExpFileFullName(Optional ByVal m_name As String) As String
    ModifiedExpFileFullName = CommComps.PendingReleaseModExpFile(m_name)
End Property

Friend Property Get RawExpFileFullName(Optional ByVal m_name As String) As String
    RawExpFileFullName = CommComps.LastModExpFileFullName(m_name)
End Property

Friend Property Let ModifiedExpFileFullName(Optional ByVal m_name As String, _
                                               ByVal m_mod_exp_file As String)
    CommComps.PendingReleaseModExpFile(m_name) = m_mod_exp_file
End Property

Friend Property Get ModifiedOnMachine(Optional ByVal m_name As String) As String
    ModifiedOnMachine = CommComps.PendingReleaseModOnMachine(m_name)
End Property

Friend Property Let ModifiedOnMachine(Optional ByVal m_name As String, _
                                        ByVal m_mod_on As String)
    CommComps.PendingReleaseModOnMachine(m_name) = m_mod_on
End Property

