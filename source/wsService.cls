VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Property Get CommonComponentsOutdated() As Long:                 CommonComponentsOutdated = Me.Range("celCommonComponentsOutdated").Value:               End Property

Public Property Let CommonComponentsOutdated(ByVal l As Long):          Me.Range("celCommonComponentsOutdated").Value = l:                                      End Property

Public Property Get CommonComponentsUpdated() As Long:                  CommonComponentsUpdated = Me.Range("celCommonComponentsUpdated").Value:                 End Property

Public Property Let CommonComponentsUpdated(ByVal l As Long):           Me.Range("celCommonComponentsUpdated").Value = l:                                       End Property

Public Property Get CommonComponentsUsed() As Long:                     CommonComponentsUsed = Me.Range("celCommonComponentsUsed").Value:                       End Property

Public Property Let CommonComponentsUsed(ByVal l As Long):              Me.Range("celCommonComponentsUsed").Value = l:                                          End Property

Public Property Get CurrentServicedWorkbookFullName() As String:        CurrentServicedWorkbookFullName = Me.Range("celCurrentServicedWorkbookFullName").Value: End Property

Public Property Let CurrentServicedWorkbookFullName(ByVal s As String): Me.Range("celCurrentServicedWorkbookFullName").Value = s:                               End Property

Public Property Get CurrentServiceLogFileFullName() As String:          CurrentServiceLogFileFullName = Me.Range("celCurrentServiceLogFileFullName").Value:     End Property

Public Property Let CurrentServiceLogFileFullName(ByVal s As String):   Me.Range("celCurrentServiceLogFileFullName").Value = s:                                 End Property

Public Property Get CurrentServiceName() As String:                     CurrentServiceName = Me.Range("celCurrentServiceName").Value:                           End Property

Public Property Let CurrentServiceName(ByVal s As String):              Me.Range("celCurrentServiceName").Value = s:                                            End Property

Public Property Get MaxLenSheetSyncId() As Long:                        MaxLenSheetSyncId = Me.Range("celMaxLenSheetSyncId").Value:                             End Property

Public Property Let MaxLenSheetSyncId(ByVal l As Long):                 Me.Range("celMaxLenSheetSyncId").Value = l:                                             End Property

Public Property Get MaxLenNameSyncId() As Long:                         MaxLenNameSyncId = Me.Range("celMaxLenNameSyncId").Value:                               End Property

Public Property Let MaxLenNameSyncId(ByVal l As Long):                  Me.Range("celMaxLenNameSyncId").Value = l:                                              End Property

Public Property Get MaxLenShapeSyncId() As Long:                        MaxLenShapeSyncId = Me.Range("celMaxLenShapeSyncId").Value:                             End Property

Public Property Let MaxLenShapeSyncId(ByVal l As Long):                 Me.Range("celMaxLenShapeSyncId").Value = l:                                             End Property

Public Property Get MaxLenVbcName() As Long:                            MaxLenVbcName = Me.Range("celMaxLenVbcName").Value:                                     End Property

Public Property Let MaxLenVbcName(ByVal l As Long):                     Me.Range("celMaxLenVbcName").Value = l:                                                 End Property

Public Property Get ServicedItemsMaxLenName() As Long:                  ServicedItemsMaxLenName = Me.Range("celServicedItemsMaxLenName").Value:                 End Property

Public Property Let ServicedItemsMaxLenName(ByVal l As Long):           Me.Range("celServicedItemsMaxLenName").Value = l:                                       End Property

Public Property Get ServicedItemsMaxLenType() As Long:                  ServicedItemsMaxLenType = Me.Range("celServicedItemsMaxLenType").Value:                 End Property

Public Property Let ServicedItemsMaxLenType(ByVal l As Long):           Me.Range("celServicedItemsMaxLenType").Value = l:                                       End Property

Public Property Get SyncComponents() As Boolean:                        SyncComponents = CBool(Me.Range("celSyncComponents").Value):                            End Property

Public Property Let SyncComponents(ByVal b As Boolean):                 Me.Range("celSyncComponents").Value = BtoYN(b):                                         End Property

Public Property Get SyncDialogLeft() As Single:                         SyncDialogLeft = Me.Range("celSyncDialogLeft").Value:                                   End Property

Public Property Let SyncDialogLeft(ByVal si As Single):                 Me.Range("celSyncDialogLeft").Value = si:                                               End Property

Public Property Get SyncDialogTitle() As String:                        SyncDialogTitle = Me.Range("celSyncDialogTitle").Value:                                 End Property

Public Property Let SyncDialogTitle(ByVal s As String):                 Me.Range("celSyncDialogTitle").Value = s:                                               End Property

Public Property Get SyncDialogTop() As Single:                          SyncDialogTop = Me.Range("celSyncDialogTop").Value:                                     End Property

Public Property Let SyncDialogTop(ByVal si As Single):                  Me.Range("celSyncDialogTop").Value = si:                                                End Property

Public Property Get SyncNames() As Boolean:                             SyncNames = CBool(Me.Range("celSyncNames").Value):                                      End Property

Public Property Let SyncNames(ByVal b As Boolean):                      Me.Range("celSyncNames").Value = CLng(b):                                               End Property

Public Property Get SyncReferences() As Boolean:                        SyncReferences = CBool(Me.Range("celSyncReferences").Value):                            End Property

Public Property Let SyncReferences(ByVal b As Boolean):                 Me.Range("celSyncReferences").Value = BtoYN(b):                                         End Property

Public Property Get SyncShapes() As Boolean:                            SyncShapes = CBool(Me.Range("celSyncShapes").Value):                                    End Property

Public Property Let SyncShapes(ByVal b As Boolean):                     Me.Range("celSyncShapes").Value = BtoYN(b):                                             End Property

Public Property Get SyncSourceFullName() As String:                     SyncSourceFullName = Me.Range("celSyncSourceFullName").Value:                           End Property

Public Property Let SyncSourceFullName(ByVal s As String):              Me.Range("celSyncSourceFullName").Value = s:                                            End Property

Public Property Get SyncTargetFullNameCopy() As String:                 SyncTargetFullNameCopy = Me.Range("celSyncTargetFullNameCopy").Value:                   End Property

Public Property Let SyncTargetFullNameCopy(ByVal s As String):          Me.Range("celSyncTargetFullNameCopy").Value = s:                                        End Property

Public Property Get SyncWorksheets() As Boolean:                        SyncWorksheets = CBool(Me.Range("celSyncWorksheets").Value):                            End Property

Public Property Let SyncWorksheets(ByVal b As Boolean):                 Me.Range("celSyncWorksheets").Value = BtoYN(b):                                         End Property

Private Function BtoYN(ByVal b As Boolean) As String
    If b Then BtoYN = "Yes" Else BtoYN = "No"
End Function
Public Sub ClearDataAllServices()
    With Me
        .CurrentServicedWorkbookFullName = vbNullString
        .ServicedItemsMaxLenName = 0
        .ServicedItemsMaxLenType = 0
    End With
End Sub

Public Sub ClearDataSynchService()
    With Me
        .MaxLenNameSyncId = 0
        .MaxLenSheetSyncId = 0
        .MaxLenVbcName = 0
    End With
End Sub

Public Sub ClearDataUpdateService()
    With Me
        .CommonComponentsUsed = 0
        .CommonComponentsOutdated = 0
        .CommonComponentsUpdated = 0
        .MaxLenVbcName = 0
    End With
End Sub

Public Sub ClearExport()

End Sub
