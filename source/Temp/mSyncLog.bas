Attribute VB_Name = "mSyncLog"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mSyncLog: Writes synchronization log records to a dedicated
' ========================  log-file.
' Public services:
' ----------------
' -
'
' ----------------------------------------------------------------------------
Private lMaxLenItem           As Long
Private lMaxLenItemType       As Long
Private lMaxLenSyncKind       As Long
Private lMaxLenSyncDone       As Long
Private lMaxLenSyncDetails    As Long

Private fso                 As New FileSystemObject

Private Property Get LogFileName() As String
    LogFileName = mSync.Target.Path & "\Sync.log"
End Property

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mSyncComps." & s
End Function

Private Function MaxLenName() As Long
'    MaxLenName = mSyncNames.MaxLenNameId(mSync.source, mSync.Target)
End Function

Private Sub WriteHeader()
' ----------------------------------------------------------------------------
' Deletes an existing log file and writes a new one with a header.
' ----------------------------------------------------------------------------
    Dim s   As String
    
    With fso
        If .FileExists(LogFileName) Then .DeleteFile (LogFileName)
    End With
    lMaxLenSyncKind = 8
    lMaxLenItemType = mBasic.Max(Len("Name"), Len("Reference"), Len("File"), Len("Worksheet"), Len("VBComponent"))
'    lMaxLenItem = mBasic.Max(mSyncNames.MaxLenNameId, mSyncRefs.MaxLenRefId, mSyncSheets.MaxLenSheetId, mSyncComps.MaxLenCompId)

    
    s = mBasic.Align("Property", lMaxLenSyncKind, , " ") & "|" & _
        mBasic.Align("Type", lMaxLenItemType, AlignCentered, " ") & "|" & _
        mBasic.Align("Item", lMaxLenItem, AlignCentered, " ") & "|" & _
        mBasic.Align("Details", lMaxLenSyncDetails, , " ")
    mFso.FileTxt(LogFileName) = s
    s = String(lMaxLenSyncKind + 1, "-") & "+" & _
        String(lMaxLenItemType + 2, "-") & "+" & _
        String(lMaxLenItem + 1, "-") & "+" & _
        String(lMaxLenSyncDetails, "-")
    mFso.FileTxt(LogFileName) = s

End Sub

Public Sub AddEntry(ByVal a_property As String, _
                    ByVal a_type As String, _
                    ByVal a_item As String, _
                    ByVal a_done As String, _
           Optional ByVal a_details As String = vbNullString)
' ----------------------------------------------------------------------------
' Writes an entry to the log file.
' ----------------------------------------------------------------------------
    Dim s As String
    
    s = mBasic.Align(a_property, lMaxLenSyncKind, , " ") & "|" & _
        mBasic.Align(a_type, lMaxLenItemType, AlignCentered, " ") & "|" & _
        mBasic.Align(a_item, lMaxLenItem, AlignCentered, " ") & "|" & _
        mBasic.Align(a_details, lMaxLenSyncDetails, , " ")
    mFso.FileTxt(LogFileName) = s

End Sub
                    

