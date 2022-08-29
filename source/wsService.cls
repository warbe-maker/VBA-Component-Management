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

Public Property Get MaxLenVbcName() As String:                          MaxLenVbcName = Me.Range("celMaxLenVbcName").Value:                             End Property

Public Property Let MaxLenVbcName(ByVal s As String):                   Me.Range("celMaxLenVbcName").Value = s:                                         End Property

Public Property Get MaxLenShapeName() As String:                        MaxLenShapeName = Me.Range("celMaxLenShapeName").Value:                         End Property

Public Property Let MaxLenShapeName(ByVal s As String):                 Me.Range("celMaxLenShapeName").Value = s:                                       End Property

Public Property Get MaxLenName() As String:                             MaxLenName = Me.Range("celMaxLenName").Value:                                   End Property

Public Property Let MaxLenName(ByVal s As String):                      Me.Range("celMaxLenName").Value = s:                                            End Property

Public Property Get MaxLenRefTo() As String:                            MaxLenRefTo = Me.Range("celMaxLenRefTo").Value:                                 End Property

Public Property Let MaxLenRefTo(ByVal s As String):                     Me.Range("celMaxLenRefTo").Value = s:                                           End Property

Public Property Get CurrentServiceName() As String:                     CurrentServiceName = Me.Range("celCurrentServiceName").Value:                   End Property

Public Property Let CurrentServiceName(ByVal s As String):              Me.Range("celCurrentServiceName").Value = s:                                    End Property

Public Property Get ServicedItemsMaxLenName() As Long:                  ServicedItemsMaxLenName = Me.Range("celServicedItemsMaxLenName").Value:         End Property

Public Property Let ServicedItemsMaxLenName(ByVal l As Long):           Me.Range("celServicedItemsMaxLenName").Value = l:                               End Property

Public Property Get ServicedItemsMaxLenType() As Long:                  ServicedItemsMaxLenType = Me.Range("celServicedItemsMaxLenType").Value:         End Property

Public Property Let ServicedItemsMaxLenType(ByVal l As Long):           Me.Range("celServicedItemsMaxLenType").Value = l:                               End Property

Public Property Get SyncDialogLeft() As Single:                         SyncDialogLeft = Me.Range("celSyncDialogLeft").Value:                           End Property

Public Property Let SyncDialogLeft(ByVal si As Single):                 Me.Range("celSyncDialogLeft").Value = si:                                       End Property

Public Property Get SyncDialogTitle() As String:                        SyncDialogTitle = Me.Range("celSyncDialogTitle").Value:                         End Property

Public Property Let SyncDialogTitle(ByVal s As String):                 Me.Range("celSyncDialogTitle").Value = s:                                       End Property

Public Property Get SyncDialogTop() As Single:                          SyncDialogTop = Me.Range("celSyncDialogTop").Value:                             End Property

Public Property Let SyncDialogTop(ByVal si As Single):                  Me.Range("celSyncDialogTop").Value = si:                                        End Property

Public Property Get SyncSourceWorkbookName() As String:                 SyncSourceWorkbookName = Me.Range("celSyncSourceWorkbookName").Value:           End Property

Public Property Let SyncSourceWorkbookName(ByVal s As String):          Me.Range("celSyncSourceWorkbookName").Value = s:                                End Property

Public Property Get SyncTargetOrigin() As String:                       SyncTargetOrigin = Me.Range("celSyncTargetOriginWorkbookFullName").Value:       End Property

Public Property Let SyncTargetOrigin(ByVal s As String):                Me.Range("celSyncTargetOriginWorkbookFullName").Value = s:                      End Property

Public Property Get SyncTargetWorkbookName() As String:                 SyncTargetWorkbookName = Me.Range("celSyncTargetWorkbookName").Value:           End Property

Public Property Let SyncTargetWorkbookName(ByVal s As String):          Me.Range("celSyncTargetWorkbookName").Value = s:                                End Property
