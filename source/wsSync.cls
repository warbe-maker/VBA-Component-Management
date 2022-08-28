VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Property Get KindColDone() As Range:                    Set KindColDone = Me.Range("colKindDone").EntireColumn:                             End Property
Private Property Get ItemsColAll() As Range:                    Set ItemsColAll = Me.Range("colItemsAll").EntireColumn:                             End Property
Private Property Get ItemsColChanged() As Range:                Set ItemsColChanged = Me.Range("colItemsChanged").EntireColumn:                     End Property
Private Property Get ItemsColChangedDone() As Range:            Set ItemsColChangedDone = Me.Range("colItemsChangedDone").EntireColumn:             End Property
Private Property Get ItemsColChangedFailed() As Range:          Set ItemsColChangedFailed = Me.Range("colItemsChangedFailed").EntireColumn:         End Property
Private Property Get ItemsColNew() As Range:                    Set ItemsColNew = Me.Range("colItemsNew").EntireColumn:                             End Property
Private Property Get ItemsColNewDone() As Range:                Set ItemsColNewDone = Me.Range("colItemsNewDone").EntireColumn:                     End Property
Private Property Get ItemsColNewFailed() As Range:              Set ItemsColNewFailed = Me.Range("colItemsNewFailed").EntireColumn:                 End Property
Private Property Get ItemsColObsolete() As Range:               Set ItemsColObsolete = Me.Range("colItemsObsolete").EntireColumn:                   End Property
Private Property Get ItemsColObsoleteDone() As Range:           Set ItemsColObsoleteDone = Me.Range("colItemsObsoleteDone").EntireColumn:           End Property
Private Property Get ItemsColObsoleteFailed() As Range:         Set ItemsColObsoleteFailed = Me.Range("colItemsObsoleteFailed").EntireColumn:       End Property
Public Property Get ItemsColPrprtys() As Range:                 Set ItemsColPrprtys = Me.Range("colShpPrprtys").EntireColumn:                       End Property
Public Property Get ItemsColPrprtysSyncExamples() As Range:     Set ItemsColPrprtysSyncExamples = Me.Range("colShpPrprtysSyncExamples").EntireColumn: End Property

Private Property Get KindOfSynchedItems() As Range:             Set KindOfSynchedItems = Me.Range("colSynchedItem").EntireColumn:                   End Property
Private Property Get NumbersAndNamesCols() As Range:            Set NumbersAndNamesCols = Me.Range("colsNumbersAndNames").EntireColumn:             End Property

Public Property Let NmeSyncDone(ByVal b As Boolean):            Intersect(KindColDone, NmeNumberRow).Value = CBoolString(b):                        End Property
Public Property Get NmeSyncDone() As Boolean:                   NmeSyncDone = CStringBool(Intersect(KindColDone, NmeNumberRow).Value):              End Property
Private Property Get NmeNameChanged() As Range:                 Set NmeNameChanged = Intersect(NmeNameTopRow, ItemsColChanged):                     End Property
Private Property Get NmeNameChangedDone() As Range:             Set NmeNameChangedDone = Intersect(NmeNameTopRow, ItemsColChangedDone):             End Property
Private Property Get NmeNameChangedFailed() As Range:           Set NmeNameChangedFailed = Intersect(NmeNameTopRow, ItemsColChangedFailed):         End Property
Private Property Get NmeNameNew() As Range:                     Set NmeNameNew = Intersect(NmeNameTopRow, ItemsColNew):                             End Property
Private Property Get NmeNameNewDone() As Range:                 Set NmeNameNewDone = Intersect(NmeNameTopRow, ItemsColNewDone):                     End Property
Private Property Get NmeNameNewFailed() As Range:               Set NmeNameNewFailed = Intersect(NmeNameTopRow, ItemsColNewFailed):                 End Property
Private Property Get NmeNameObsolete() As Range:                Set NmeNameObsolete = Intersect(NmeNameTopRow, ItemsColObsolete):                   End Property
Private Property Get NmeNameObsoleteDone() As Range:            Set NmeNameObsoleteDone = Intersect(NmeNameTopRow, ItemsColObsoleteDone):           End Property
Private Property Get NmeNameObsoleteFailed() As Range:          Set NmeNameObsoleteFailed = Intersect(NmeNameTopRow, ItemsColObsoleteFailed):       End Property
Private Property Get NmeNameTopRow() As Range:                  Set NmeNameTopRow = NmeNumberRow.Offset(1, 0).EntireRow:                            End Property
Private Property Get NmeNumberRow() As Range:                   Set NmeNumberRow = Me.Range("rowNmeFigures").EntireRow:                             End Property
Private Property Let NmeNumberAll(ByVal l As Long):             Intersect(ItemsColAll, NmeNumberRow).Value = l:                                     End Property
Public Property Get NmeNumberAll() As Long:                     NmeNumberAll = Intersect(ItemsColAll, NmeNumberRow).Value:                          End Property
Public Property Get NmeNumberChanged() As Long:                 NmeNumberChanged = Intersect(ItemsColChanged, NmeNumberRow).Value:                  End Property
Public Property Get NmeNumberChangedDone() As Long:             NmeNumberChangedDone = Intersect(ItemsColChangedDone, NmeNumberRow).Value:          End Property
Public Property Get NmeNumberChangedFailed() As Long:           NmeNumberChangedFailed = Intersect(ItemsColChangedFailed, NmeNumberRow).Value:      End Property
Public Property Get NmeNumberNew() As Long:                     NmeNumberNew = Intersect(ItemsColNew, NmeNumberRow).Value:                          End Property
Public Property Get NmeNumberNewDone() As Long:                 NmeNumberNewDone = Intersect(ItemsColNewDone, NmeNumberRow).Value:                  End Property
Public Property Get NmeNumberNewFailed() As Long:               NmeNumberNewFailed = Intersect(ItemsColNewFailed, NmeNumberRow).Value:              End Property
Public Property Get NmeNumberObsolete() As Long:                NmeNumberObsolete = Intersect(ItemsColObsolete, NmeNumberRow).Value:                End Property
Public Property Get NmeNumberObsoleteDone() As Long:            NmeNumberObsoleteDone = Intersect(ItemsColObsoleteDone, NmeNumberRow).Value:        End Property
Public Property Get NmeNumberObsoleteFailed() As Long:          NmeNumberObsoleteFailed = Intersect(ItemsColObsoleteFailed, NmeNumberRow).Value:    End Property
Public Property Let NmeNumberChanged(ByVal l As Long):          Intersect(ItemsColChanged, NmeNumberRow).Value = l:                                 End Property
Public Property Let NmeNumberChangedDone(ByVal l As Long):      Intersect(ItemsColChangedDone, NmeNumberRow).Value = l:                             End Property
Public Property Let NmeNumberChangedFailed(ByVal l As Long):    Intersect(ItemsColChangedFailed, NmeNumberRow).Value = l:                           End Property
Public Property Let NmeNumberNew(ByVal l As Long):              Intersect(ItemsColNew, NmeNumberRow).Value = l:                                     End Property
Public Property Let NmeNumberNewDone(ByVal l As Long):          Intersect(ItemsColNewDone, NmeNumberRow).Value = l:                                 End Property
Public Property Let NmeNumberNewFailed(ByVal l As Long):        Intersect(ItemsColNewFailed, NmeNumberRow).Value = l:                               End Property
Public Property Let NmeNumberObsolete(ByVal l As Long):         Intersect(ItemsColObsolete, NmeNumberRow).Value = l:                                End Property
Public Property Let NmeNumberObsoleteDone(ByVal l As Long):     Intersect(ItemsColObsoleteDone, NmeNumberRow).Value = l:                            End Property
Public Property Let NmeNumberObsoleteFailed(ByVal l As Long):   Intersect(ItemsColObsoleteFailed, NmeNumberRow).Value = l:                          End Property

Public Property Let RefSyncDone(ByVal b As Boolean):            Intersect(KindColDone, RefNumberRow).Value = CBoolString(b):                          End Property
Public Property Get RefSyncDone() As Boolean:                   RefSyncDone = CStringBool(Intersect(KindColDone, RefNumberRow).Value):                    End Property
Private Property Get RefNameNew() As Range:                     Set RefNameNew = Intersect(RefNameTopRow, ItemsColNew):                             End Property
Private Property Get RefNameNewDone() As Range:                 Set RefNameNewDone = Intersect(RefNameTopRow, ItemsColNewDone):                     End Property
Private Property Get RefNameNewFailed() As Range:               Set RefNameNewFailed = Intersect(RefNameTopRow, ItemsColNewFailed):                 End Property
Private Property Get RefNameObsolete() As Range:                Set RefNameObsolete = Intersect(RefNameTopRow, ItemsColObsolete):                   End Property
Private Property Get RefNameObsoleteDone() As Range:            Set RefNameObsoleteDone = Intersect(RefNameTopRow, ItemsColObsoleteDone):           End Property
Private Property Get RefNameObsoleteFailed() As Range:          Set RefNameObsoleteFailed = Intersect(RefNameTopRow, ItemsColObsoleteFailed):       End Property
Private Property Get RefNameTopRow() As Range:                  Set RefNameTopRow = RefNumberRow.Offset(1, 0).EntireRow:                            End Property
Private Property Get RefNumberRow() As Range:                   Set RefNumberRow = Me.Range("rowRefFigures").EntireRow:                             End Property
Private Property Get RefNumberAll() As Long:                    RefNumberAll = Intersect(ItemsColAll, RefNumberRow).Value:                          End Property
Private Property Let RefNumberAll(ByVal l As Long):             Intersect(ItemsColAll, RefNumberRow).Value = l:                                     End Property
Public Property Get RefNumberNew() As Long:                     RefNumberNew = Intersect(ItemsColNew, RefNumberRow).Value:                          End Property
Public Property Get RefNumberNewDone() As Long:                 RefNumberNewDone = Intersect(ItemsColNewDone, RefNumberRow).Value:                  End Property
Public Property Get RefNumberNewFailed() As Long:               RefNumberNewFailed = Intersect(ItemsColNewFailed, RefNumberRow).Value:              End Property
Public Property Get RefNumberObsolete() As Long:                RefNumberObsolete = Intersect(ItemsColObsolete, RefNumberRow).Value:                End Property
Public Property Get RefNumberObsoleteDone() As Long:            RefNumberObsoleteDone = Intersect(ItemsColObsoleteDone, RefNumberRow).Value:        End Property
Public Property Get RefNumberObsoleteFailed() As Long:          RefNumberObsoleteFailed = Intersect(ItemsColObsoleteFailed, RefNumberRow).Value:    End Property
Public Property Let RefNumberNew(ByVal l As Long):              Intersect(ItemsColNew, RefNumberRow).Value = l:                                     End Property
Public Property Let RefNumberNewDone(ByVal l As Long):          Intersect(ItemsColNewDone, RefNumberRow).Value = l:                                 End Property
Public Property Let RefNumberNewFailed(ByVal l As Long):        Intersect(ItemsColNewFailed, RefNumberRow).Value = l:                               End Property
Public Property Let RefNumberObsolete(ByVal l As Long):         Intersect(ItemsColObsolete, RefNumberRow).Value = l:                                End Property
Public Property Let RefNumberObsoleteDone(ByVal l As Long):     Intersect(ItemsColObsoleteDone, RefNumberRow).Value = l:                            End Property
Public Property Let RefNumberObsoleteFailed(ByVal l As Long):   Intersect(ItemsColObsoleteFailed, RefNumberRow).Value = l:                          End Property

Public Property Let RngSyncDone(ByVal b As Boolean):            Intersect(KindColDone, RngNumberRow).Value = CBoolString(b):                          End Property
Public Property Get RngSyncDone() As Boolean:                   RngSyncDone = CStringBool(Intersect(KindColDone, RngNumberRow).Value):                    End Property
Private Property Get RngNameChanged() As Range:                 Set RngNameChanged = Intersect(RngNameTopRow, ItemsColChanged):                     End Property
Private Property Get RngNameChangedDone() As Range:             Set RngNameChangedDone = Intersect(RngNameTopRow, ItemsColChangedDone):             End Property
Private Property Get RngNameChangedFailed() As Range:           Set RngNameChangedFailed = Intersect(RngNameTopRow, ItemsColChangedFailed):         End Property
Private Property Get RngNameNew() As Range:                     Set RngNameNew = Intersect(RngNameTopRow, ItemsColNew):                             End Property
Private Property Get RngNameNewDone() As Range:                 Set RngNameNewDone = Intersect(RngNameTopRow, ItemsColNewDone):                     End Property
Private Property Get RngNameNewFailed() As Range:               Set RngNameNewFailed = Intersect(RngNameTopRow, ItemsColNewFailed):                 End Property
Private Property Get RngNameObsolete() As Range:                Set RngNameObsolete = Intersect(RngNameTopRow, ItemsColObsolete):                   End Property
Private Property Get RngNameObsoleteDone() As Range:            Set RngNameObsoleteDone = Intersect(RngNameTopRow, ItemsColObsoleteDone):           End Property
Private Property Get RngNameObsoleteFailed() As Range:          Set RngNameObsoleteFailed = Intersect(RngNameTopRow, ItemsColObsoleteFailed):       End Property
Private Property Get RngNameTopRow() As Range:                  Set RngNameTopRow = RngNumberRow.Offset(1, 0).EntireRow:                            End Property
Private Property Get RngNumberAll() As Long:                    RngNumberAll = Intersect(ItemsColAll, RngNumberRow).Value:                          End Property
Private Property Get RngNumberRow() As Range:                   Set RngNumberRow = Me.Range("rowRngFigures").EntireRow:                             End Property
Private Property Let RngNumberAll(ByVal l As Long):             Intersect(ItemsColAll, RngNumberRow).Value = l:                                     End Property
Public Property Get RngNumberChanged() As Long:                 RngNumberChanged = Intersect(ItemsColChanged, RngNumberRow).Value:                  End Property
Public Property Get RngNumberChangedDone() As Long:             RngNumberChangedDone = Intersect(ItemsColChangedDone, RngNumberRow).Value:          End Property
Public Property Get RngNumberChangedFailed() As Long:           RngNumberChangedFailed = Intersect(ItemsColChangedFailed, RngNumberRow).Value:      End Property
Public Property Get RngNumberNew() As Long:                     RngNumberNew = Intersect(ItemsColNew, RngNumberRow).Value:                          End Property
Public Property Get RngNumberNewDone() As Long:                 RngNumberNewDone = Intersect(ItemsColNewDone, RngNumberRow).Value:                  End Property
Public Property Get RngNumberNewFailed() As Long:               RngNumberNewFailed = Intersect(ItemsColNewFailed, RngNumberRow).Value:              End Property
Public Property Get RngNumberObsolete() As Long:                RngNumberObsolete = Intersect(ItemsColObsolete, RngNumberRow).Value:                End Property
Public Property Get RngNumberObsoleteDone() As Long:            RngNumberObsoleteDone = Intersect(ItemsColObsoleteDone, RngNumberRow).Value:        End Property
Public Property Get RngNumberObsoleteFailed() As Long:          RngNumberObsoleteFailed = Intersect(ItemsColObsoleteFailed, RngNumberRow).Value:    End Property
Public Property Let RngNumberChanged(ByVal l As Long):          Intersect(ItemsColChanged, RngNumberRow).Value = l:                                 End Property
Public Property Let RngNumberChangedDone(ByVal l As Long):      Intersect(ItemsColChangedDone, RngNumberRow).Value = l:                             End Property
Public Property Let RngNumberChangedFailed(ByVal l As Long):    Intersect(ItemsColChangedFailed, RngNumberRow).Value = l:                           End Property
Public Property Let RngNumberNew(ByVal l As Long):              Intersect(ItemsColNew, RngNumberRow).Value = l:                                     End Property
Public Property Let RngNumberNewDone(ByVal l As Long):          Intersect(ItemsColNewDone, RngNumberRow).Value = l:                                 End Property
Public Property Let RngNumberNewFailed(ByVal l As Long):        Intersect(ItemsColNewFailed, RngNumberRow).Value = l:                               End Property
Public Property Let RngNumberObsolete(ByVal l As Long):         Intersect(ItemsColObsolete, RngNumberRow).Value = l:                                End Property
Public Property Let RngNumberObsoleteDone(ByVal l As Long):     Intersect(ItemsColObsoleteDone, RngNumberRow).Value = l:                            End Property
Public Property Let RngNumberObsoleteFailed(ByVal l As Long):   Intersect(ItemsColObsoleteFailed, RngNumberRow).Value = l:                          End Property

Public Property Let ShpSyncDone(ByVal b As Boolean):            Intersect(KindColDone, ShpNumberRow).Value = CBoolString(b):                          End Property
Public Property Get ShpSyncDone() As Boolean:                   ShpSyncDone = CStringBool(Intersect(KindColDone, ShpNumberRow).Value):                    End Property
Private Property Get ShpNameChanged() As Range:                 Set ShpNameChanged = Intersect(ShpNameTopRow, ItemsColChanged):                     End Property
Private Property Get ShpNameChangedDone() As Range:             Set ShpNameChangedDone = Intersect(ShpNameTopRow, ItemsColChangedDone):             End Property
Private Property Get ShpNameChangedFailed() As Range:           Set ShpNameChangedFailed = Intersect(ShpNameTopRow, ItemsColChangedFailed):         End Property
Private Property Get ShpNameNew() As Range:                     Set ShpNameNew = Intersect(ShpNameTopRow, ItemsColNew):                             End Property
Private Property Get ShpNameNewDone() As Range:                 Set ShpNameNewDone = Intersect(ShpNameTopRow, ItemsColNewDone):                     End Property
Private Property Get ShpNameNewFailed() As Range:               Set ShpNameNewFailed = Intersect(ShpNameTopRow, ItemsColNewFailed):                 End Property
Private Property Get ShpNameObsolete() As Range:                Set ShpNameObsolete = Intersect(ShpNameTopRow, ItemsColObsolete):                   End Property
Private Property Get ShpNameObsoleteDone() As Range:            Set ShpNameObsoleteDone = Intersect(ShpNameTopRow, ItemsColObsoleteDone):           End Property
Private Property Get ShpNameObsoleteFailed() As Range:          Set ShpNameObsoleteFailed = Intersect(ShpNameTopRow, ItemsColObsoleteFailed):       End Property
Private Property Get ShpNamePrprtys() As Range:                 Set ShpNamePrprtys = Intersect(ShpNameTopRow, ItemsColPrprtys):                     End Property
Private Property Get ShpNamePrprtysSyncExamples() As Range:     Set ShpNamePrprtysSyncExamples = Intersect(ShpNameTopRow, ItemsColPrprtysSyncExamples):  End Property
Private Property Get ShpNameTopRow() As Range:                  Set ShpNameTopRow = ShpNumberRow.Offset(1, 0).EntireRow:                            End Property
Private Property Let ShpNumberAll(ByVal l As Long):             Me.Range("celShpItemsTotal").Value = l:                                             End Property
Public Property Get ShpNumberChanged() As Long:                 ShpNumberChanged = Me.Range("celShpItemsChanged").Value:                            End Property
Public Property Get ShpNumberChangedDone() As Long:             ShpNumberChangedDone = Me.Range("celShpItemsChangedDone").Value:                    End Property
Public Property Get ShpNumberChangedFailed() As Long:           ShpNumberChangedFailed = Me.Range("celShpItemsChangedFailed").Value:                End Property
Public Property Get ShpNumberNew() As Long:                     ShpNumberNew = Me.Range("celShpItemsNew").Value:                                    End Property
Public Property Get ShpNumberNewDone() As Long:                 ShpNumberNewDone = Me.Range("celShpItemsNewDone").Value:                            End Property
Public Property Get ShpNumberNewFailed() As Long:               ShpNumberNewFailed = Me.Range("celShpItemsNewFailed").Value:                        End Property
Public Property Get ShpNumberObsolete() As Long:                ShpNumberObsolete = Me.Range("celShpItemsObsolete").Value:                          End Property
Public Property Get ShpNumberObsoleteDone() As Long:            ShpNumberObsoleteDone = Me.Range("celShpItemsObsoleteDone").Value:                  End Property
Public Property Get ShpNumberObsoleteFailed() As Long:          ShpNumberObsoleteFailed = Me.Range("celShpItemsObsoleteFailed").Value:              End Property
Public Property Get ShpNumberPrprtys() As Long:                 ShpNumberPrprtys = Intersect(ItemsColPrprtys, ShpNumberRow).Value:                  End Property
Public Property Get ShpNumberPrprtysSyncExamples() As Long:     ShpNumberPrprtysSyncExamples = Intersect(ItemsColPrprtysSyncExamples, ShpNumberRow).Value:   End Property
Public Property Get ShpNumberRow() As Range:                    Set ShpNumberRow = Me.Range("rowShpFigures").EntireRow:                             End Property
Public Property Get ShpNumberAll() As Long:                     ShpNumberAll = Me.Range("celShpItemsTotal").Value:                                  End Property
Public Property Let ShpNumberChanged(ByVal l As Long):          Me.Range("celShpItemsChanged").Value = l:                                           End Property
Public Property Let ShpNumberChangedDone(ByVal l As Long):      Me.Range("celShpItemsChangedDone").Value = l:                                       End Property
Public Property Let ShpNumberChangedFailed(ByVal l As Long):    Me.Range("celShpItemsChangedFailed").Value = l:                                     End Property
Public Property Let ShpNumberNew(ByVal l As Long):              Me.Range("celShpItemsNew").Value = l:                                               End Property
Public Property Let ShpNumberNewDone(ByVal l As Long):          Me.Range("celShpItemsNewDone").Value = l:                                           End Property
Public Property Let ShpNumberNewFailed(ByVal l As Long):        Me.Range("celShpItemsNewFailed").Value = l:                                         End Property
Public Property Let ShpNumberObsolete(ByVal l As Long):         Me.Range("celShpItemsObsolete").Value = l:                                          End Property
Public Property Let ShpNumberObsoleteDone(ByVal l As Long):     Me.Range("celShpItemsObsoleteDone").Value = l:                                      End Property
Public Property Let ShpNumberObsoleteFailed(ByVal l As Long):   Me.Range("celShpItemsObsoleteFailed").Value = l:                                    End Property
Public Property Let ShpNumberPrprtys(ByVal l As Long):          Intersect(ItemsColPrprtys, ShpNumberRow).Value = l:                                 End Property
Public Property Let ShpNumberPrprtysSyncExamples(ByVal l As Long): Intersect(ItemsColPrprtysSyncExamples, ShpNumberRow).Value = l:                  End Property

Public Property Let VbcSyncDone(ByVal b As Boolean):            Intersect(KindColDone, VbcNumberRow).Value = CBoolString(b):                        End Property
Public Property Get VbcSyncDone() As Boolean:                   VbcSyncDone = CStringBool(Intersect(KindColDone, VbcNumberRow).Value):              End Property
Private Property Get VbcNameChanged() As Range:                 Set VbcNameChanged = Intersect(VbcNameTopRow, ItemsColChanged):                     End Property
Private Property Get VbcNameChangedDone() As Range:             Set VbcNameChangedDone = Intersect(VbcNameTopRow, ItemsColChangedDone):             End Property
Private Property Get VbcNameChangedFailed() As Range:           Set VbcNameChangedFailed = Intersect(VbcNameTopRow, ItemsColChangedFailed):         End Property
Private Property Get VbcNameNew() As Range:                     Set VbcNameNew = Intersect(VbcNameTopRow, ItemsColNew):                             End Property
Private Property Get VbcNameNewDone() As Range:                 Set VbcNameNewDone = Intersect(VbcNameTopRow, ItemsColNewDone):                     End Property
Private Property Get VbcNameNewFailed() As Range:               Set VbcNameNewFailed = Intersect(VbcNameTopRow, ItemsColNewFailed):                 End Property
Private Property Get VbcNameObsolete() As Range:                Set VbcNameObsolete = Intersect(VbcNameTopRow, ItemsColObsolete):                   End Property
Private Property Get VbcNameObsoleteDone() As Range:            Set VbcNameObsoleteDone = Intersect(VbcNameTopRow, ItemsColObsoleteDone):           End Property
Private Property Get VbcNameObsoleteFailed() As Range:          Set VbcNameObsoleteFailed = Intersect(VbcNameTopRow, ItemsColObsoleteFailed):       End Property
Private Property Get VbcNameTopRow() As Range:                  Set VbcNameTopRow = VbcNumberRow.Offset(1, 0).EntireRow:                            End Property
Private Property Get VbcNumberRow() As Range:                   Set VbcNumberRow = Me.Range("rowVbcFigures").EntireRow:                             End Property
Private Property Get VbcNumberAll() As Long:                    VbcNumberAll = Intersect(ItemsColAll, VbcNumberRow).Value:                          End Property
Private Property Let VbcNumberAll(ByVal l As Long):             Intersect(ItemsColAll, VbcNumberRow).Value = l:                                     End Property
Public Property Get VbcNumberChanged() As Long:                 VbcNumberChanged = Intersect(ItemsColChanged, VbcNumberRow).Value:                  End Property
Public Property Get VbcNumberChangedDone() As Long:             VbcNumberChangedDone = Intersect(ItemsColChangedDone, VbcNumberRow).Value:          End Property
Public Property Get VbcNumberChangedFailed() As Long:           VbcNumberChangedFailed = Intersect(ItemsColChangedFailed, VbcNumberRow).Value:      End Property
Public Property Get VbcNumberNew() As Long:                     VbcNumberNew = Intersect(ItemsColNew, VbcNumberRow).Value:                          End Property
Public Property Get VbcNumberNewDone() As Long:                 VbcNumberNewDone = Intersect(ItemsColNewDone, VbcNumberRow).Value:                  End Property
Public Property Get VbcNumberNewFailed() As Long:               VbcNumberNewFailed = Intersect(ItemsColNewFailed, VbcNumberRow).Value:              End Property
Public Property Get VbcNumberObsolete() As Long:                VbcNumberObsolete = Intersect(ItemsColObsolete, VbcNumberRow).Value:                End Property
Public Property Get VbcNumberObsoleteDone() As Long:            VbcNumberObsoleteDone = Intersect(ItemsColObsoleteDone, VbcNumberRow).Value:        End Property
Public Property Get VbcNumberObsoleteFailed() As Long:          VbcNumberObsoleteFailed = Intersect(ItemsColObsoleteFailed, VbcNumberRow).Value:    End Property
Public Property Let VbcNumberChanged(ByVal l As Long):          Intersect(ItemsColChanged, VbcNumberRow).Value = l:                                 End Property
Public Property Let VbcNumberChangedDone(ByVal l As Long):      Intersect(ItemsColChangedDone, VbcNumberRow).Value = l:                             End Property
Public Property Let VbcNumberChangedFailed(ByVal l As Long):    Intersect(ItemsColChangedFailed, VbcNumberRow).Value = l:                           End Property
Public Property Let VbcNumberNew(ByVal l As Long):              Intersect(ItemsColNew, VbcNumberRow).Value = l:                                     End Property
Public Property Let VbcNumberNewDone(ByVal l As Long):          Intersect(ItemsColNewDone, VbcNumberRow).Value = l:                                 End Property
Public Property Let VbcNumberNewFailed(ByVal l As Long):        Intersect(ItemsColNewFailed, VbcNumberRow).Value = l:                               End Property
Public Property Let VbcNumberObsolete(ByVal l As Long):         Intersect(ItemsColObsolete, VbcNumberRow).Value = l:                                End Property
Public Property Let VbcNumberObsoleteDone(ByVal l As Long):     Intersect(ItemsColObsoleteDone, VbcNumberRow).Value = l:                            End Property
Public Property Let VbcNumberObsoleteFailed(ByVal l As Long):   Intersect(ItemsColObsoleteFailed, VbcNumberRow).Value = l:                          End Property

Public Property Let WshSyncDone(ByVal b As Boolean):            Intersect(KindColDone, WshNumberRow).Value = CBoolString(b):                          End Property
Public Property Get WshSyncDone() As Boolean:                   WshSyncDone = CStringBool(Intersect(KindColDone, WshNumberRow).Value):                    End Property
Private Property Get WshNameChanged() As Range:                 Set WshNameChanged = Intersect(WshNameTopRow, ItemsColChanged):                     End Property
Private Property Get WshNameChangedDone() As Range:             Set WshNameChangedDone = Intersect(WshNameTopRow, ItemsColChangedDone):             End Property
Private Property Get WshNameChangedFailed() As Range:           Set WshNameChangedFailed = Intersect(WshNameTopRow, ItemsColChangedFailed):         End Property
Private Property Get WshNameNew() As Range:                     Set WshNameNew = Intersect(WshNameTopRow, ItemsColNew):                             End Property
Private Property Get WshNameNewDone() As Range:                 Set WshNameNewDone = Intersect(WshNameTopRow, ItemsColNewDone):                     End Property
Private Property Get WshNameNewFailed() As Range:               Set WshNameNewFailed = Intersect(WshNameTopRow, ItemsColNewFailed):                 End Property
Private Property Get WshNameObsolete() As Range:                Set WshNameObsolete = Intersect(WshNameTopRow, ItemsColObsolete):                   End Property
Private Property Get WshNameObsoleteDone() As Range:            Set WshNameObsoleteDone = Intersect(WshNameTopRow, ItemsColObsoleteDone):           End Property
Private Property Get WshNameObsoleteFailed() As Range:          Set WshNameObsoleteFailed = Intersect(WshNameTopRow, ItemsColObsoleteFailed):       End Property
Private Property Get WshNameTopRow() As Range:                  Set WshNameTopRow = WshNumberRow.Offset(1, 0).EntireRow:                            End Property
Private Property Get WshNumberRow() As Range:                   Set WshNumberRow = Me.Range("rowWshFigures").EntireRow:                             End Property
Private Property Get WshNumberAll() As Long:                    WshNumberAll = Intersect(ItemsColAll, WshNumberRow).Value:                          End Property
Private Property Let WshNumberAll(ByVal l As Long):             Intersect(ItemsColAll, WshNumberRow).Value = l:                                     End Property
Public Property Get WshNumberChanged() As Long:                 WshNumberChanged = Intersect(ItemsColChanged, WshNumberRow).Value:                  End Property
Public Property Get WshNumberChangedDone() As Long:             WshNumberChangedDone = Intersect(ItemsColChangedDone, WshNumberRow).Value:          End Property
Public Property Get WshNumberChangedFailed() As Long:           WshNumberChangedFailed = Intersect(ItemsColChangedFailed, WshNumberRow).Value:      End Property
Public Property Get WshNumberNew() As Long:                     WshNumberNew = Intersect(ItemsColNew, WshNumberRow).Value:                          End Property
Public Property Get WshNumberNewDone() As Long:                 WshNumberNewDone = Intersect(ItemsColNewDone, WshNumberRow).Value:                  End Property
Public Property Get WshNumberNewFailed() As Long:               WshNumberNewFailed = Intersect(ItemsColNewFailed, WshNumberRow).Value:              End Property
Public Property Get WshNumberObsolete() As Long:                WshNumberObsolete = Intersect(ItemsColObsolete, WshNumberRow).Value:                End Property
Public Property Get WshNumberObsoleteDone() As Long:            WshNumberObsoleteDone = Intersect(ItemsColObsoleteDone, WshNumberRow).Value:        End Property
Public Property Get WshNumberObsoleteFailed() As Long:          WshNumberObsoleteFailed = Intersect(ItemsColObsoleteFailed, WshNumberRow).Value:    End Property
Public Property Let WshNumberChanged(ByVal l As Long):          Intersect(ItemsColChanged, WshNumberRow).Value = l:                                 End Property
Public Property Let WshNumberChangedDone(ByVal l As Long):      Intersect(ItemsColChangedDone, WshNumberRow).Value = l:                             End Property
Public Property Let WshNumberChangedFailed(ByVal l As Long):    Intersect(ItemsColChangedFailed, WshNumberRow).Value = l:                           End Property
Public Property Let WshNumberNew(ByVal l As Long):              Intersect(ItemsColNew, WshNumberRow).Value = l:                                     End Property
Public Property Let WshNumberNewDone(ByVal l As Long):          Intersect(ItemsColNewDone, WshNumberRow).Value = l:                                 End Property
Public Property Let WshNumberNewFailed(ByVal l As Long):        Intersect(ItemsColNewFailed, WshNumberRow).Value = l:                               End Property
Public Property Let WshNumberObsolete(ByVal l As Long):         Intersect(ItemsColObsolete, WshNumberRow).Value = l:                                End Property
Public Property Let WshNumberObsoleteDone(ByVal l As Long):     Intersect(ItemsColObsoleteDone, WshNumberRow).Value = l:                            End Property
Public Property Let WshNumberObsoleteFailed(ByVal l As Long):   Intersect(ItemsColObsoleteFailed, WshNumberRow).Value = l:                          End Property

Public Property Get NmeItemAll(Optional ByVal s As String) As Boolean
    NmeItemAll = NameFound(NmeNumberAll, Intersect(ItemsColAll, NmeNumberRow), s)
End Property
Public Property Let NmeItemAll(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColAll, NmeNameTopRow.Offset(NmeNumberAll, 0)), s
    NmeNumberAll = NmeNumberAll + 1
End Property

Public Property Get NmeItemNew(Optional ByVal s As String) As Boolean
    NmeItemNew = NameFound(NmeNumberNew, Intersect(ItemsColNew, NmeNumberRow), s)
End Property
Public Property Let NmeItemNew(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNew, NmeNameTopRow.Offset(NmeNumberNew, 0)), s
    NmeNumberNew = NmeNumberNew + 1
End Property

Public Property Get NmeItemNewDone(Optional ByVal s As String) As Boolean
    NmeItemNewDone = NameFound(NmeNumberNewDone, Intersect(ItemsColNew, NmeNumberRow), s)
End Property
Public Property Let NmeItemNewDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewDone, NmeNameNewDone.Offset(NmeNumberNewDone, 0)), s
    NmeNumberNewDone = NmeNumberNewDone + 1
End Property

Public Property Get NmeItemNewFailed(Optional ByVal s As String) As Boolean
    NmeItemNewFailed = NameFound(NmeNumberNewFailed, Intersect(ItemsColNewFailed, NmeNumberRow), s)
End Property
Public Property Let NmeItemNewFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewFailed, NmeNameNewFailed.Offset(NmeNumberNewFailed, 0)), s
    NmeNumberNewFailed = NmeNumberNewFailed + 1
End Property

Public Property Get NmeItemObsolete(Optional ByVal s As String) As Boolean
    NmeItemObsolete = NameFound(NmeNumberObsolete, Intersect(ItemsColObsolete, NmeNumberRow), s)
End Property
Public Property Let NmeItemObsolete(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsolete, NmeNameObsolete.Offset(NmeNumberObsolete, 0)), s
    NmeNumberObsolete = NmeNumberObsolete + 1
End Property

Public Property Get NmeItemObsoleteDone(Optional ByVal s As String) As Boolean
    NmeItemObsoleteDone = NameFound(NmeNumberObsoleteDone, Intersect(ItemsColObsoleteDone, NmeNumberRow), s)
End Property
Public Property Let NmeItemObsoleteDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteDone, NmeNameObsoleteDone.Offset(NmeNumberObsoleteDone, 0)), s
    NmeNumberObsoleteDone = NmeNumberObsoleteDone + 1
End Property

Public Property Get NmeItemObsoleteFailed(Optional ByVal s As String) As Boolean
    NmeItemObsoleteFailed = NameFound(NmeNumberObsoleteFailed, Intersect(ItemsColObsoleteFailed, NmeNumberRow), s)
End Property
Public Property Let NmeItemObsoleteFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteFailed, NmeNameObsoleteFailed.Offset(NmeNumberObsoleteFailed, 0)), s
    NmeNumberObsoleteFailed = NmeNumberObsoleteFailed + 1
End Property

Public Property Get NmeItemChanged(Optional ByVal s As String) As Boolean
    NmeItemChanged = NameFound(NmeNumberChanged, Intersect(ItemsColChanged, NmeNumberRow), s)
End Property
Public Property Let NmeItemChanged(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChanged, NmeNameChanged.Offset(NmeNumberChanged, 0)), s
    NmeNumberChanged = NmeNumberChanged + 1
End Property

Public Property Get NmeItemChangedDone(Optional ByVal s As String) As Boolean
    NmeItemChangedDone = NameFound(NmeNumberChangedDone, Intersect(ItemsColChangedDone, NmeNumberRow), s)
End Property
Public Property Let NmeItemChangedDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedDone, NmeNameChangedDone.Offset(NmeNumberChangedDone, 0)), s
    NmeNumberChangedDone = NmeNumberChangedDone + 1
End Property

Public Property Get NmeItemChangedFailed(Optional ByVal s As String) As Boolean
    NmeItemChangedFailed = NameFound(NmeNumberChangedFailed, Intersect(ItemsColChangedFailed, NmeNumberRow), s)
End Property
Public Property Let NmeItemChangedFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedFailed, NmeNameChangedFailed.Offset(NmeNumberChangedFailed, 0)), s
    NmeNumberChangedFailed = NmeNumberChangedFailed + 1
End Property




















Public Property Get RefItemAll(Optional ByVal s As String) As Boolean
    RefItemAll = NameFound(RefNumberAll, Intersect(ItemsColAll, RefNumberRow), s)
End Property
Public Property Let RefItemAll(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColAll, RefNameTopRow.Offset(RefNumberAll, 0)), s
    RefNumberAll = RefNumberAll + 1
End Property

Public Property Get RefItemNew(Optional ByVal s As String) As Boolean
    RefItemNew = NameFound(RefNumberNew, Intersect(ItemsColNew, RefNumberRow), s)
End Property
Public Property Let RefItemNew(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNew, RefNameTopRow.Offset(RefNumberNew, 0)), s
    RefNumberNew = RefNumberNew + 1
End Property

Public Property Get RefItemNewDone(Optional ByVal s As String) As Boolean
    RefItemNewDone = NameFound(RefNumberNewDone, Intersect(ItemsColNewDone, RefNumberRow), s)
End Property
Public Property Let RefItemNewDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewDone, RefNameNewDone.Offset(RefNumberNewDone, 0)), s
    RefNumberNewDone = RefNumberNewDone + 1
End Property

Public Property Get RefItemNewFailed(Optional ByVal s As String) As Boolean
    RefItemNewFailed = NameFound(RefNumberNewFailed, Intersect(ItemsColNewFailed, RefNumberRow), s)
End Property
Public Property Let RefItemNewFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewFailed, RefNameNewFailed.Offset(RefNumberNewFailed, 0)), s
    RefNumberNewFailed = RefNumberNewFailed + 1
End Property

Public Property Get RefItemObsolete(Optional ByVal s As String) As Boolean
    RefItemObsolete = NameFound(RefNumberObsolete, Intersect(ItemsColObsolete, RefNumberRow), s)
End Property
Public Property Let RefItemObsolete(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsolete, RefNameObsolete.Offset(RefNumberObsolete, 0)), s
    RefNumberObsolete = RefNumberObsolete + 1
End Property

Public Property Get RefItemObsoleteDone(Optional ByVal s As String) As Boolean
    RefItemObsoleteDone = NameFound(RefNumberObsoleteDone, Intersect(ItemsColObsoleteDone, RefNumberRow), s)
End Property
Public Property Let RefItemObsoleteDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteDone, RefNameObsoleteDone.Offset(RefNumberObsoleteDone, 0)), s
    RefNumberObsoleteDone = RefNumberObsoleteDone + 1
End Property

Public Property Get RefItemObsoleteFailed(Optional ByVal s As String) As Boolean
    RefItemObsoleteFailed = NameFound(RefNumberObsoleteFailed, Intersect(ItemsColObsoleteFailed, RefNumberRow), s)
End Property
Public Property Let RefItemObsoleteFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteFailed, RefNameObsoleteFailed.Offset(RefNumberObsoleteFailed, 0)), s
    RefNumberObsoleteFailed = RefNumberObsoleteFailed + 1
End Property


Public Property Get ShpItemAll(Optional ByVal s As String) As Boolean
    ShpItemAll = NameFound(ShpNumberAll, Intersect(ItemsColAll, ShpNumberRow), s)
End Property
Public Property Let ShpItemAll(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColAll, ShpNameTopRow.Offset(ShpNumberAll, 0)), s
    ShpNumberAll = ShpNumberAll + 1
End Property

Public Property Get ShpItemNew(Optional ByVal s As String) As Boolean
    ShpItemNew = NameFound(ShpNumberNew, Me.Range("celShpItemsNew").Offset(1, 0), s)
End Property
Public Property Let ShpItemNew(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNew, ShpNameTopRow.Offset(ShpNumberNew, 0)), s
    ShpNumberNew = ShpNumberNew + 1
End Property

Public Property Get ShpItemNewDone(Optional ByVal s As String) As Boolean
    ShpItemNewDone = NameFound(ShpNumberNewDone, ShpNameNewDone, s)
End Property
Public Property Let ShpItemNewDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewDone, ShpNameNewDone.Offset(ShpNumberNewDone, 0)), s
    ShpNumberNewDone = ShpNumberNewDone + 1
End Property

Public Property Get ShpItemNewFailed(Optional ByVal s As String) As Boolean
    ShpItemNewFailed = NameFound(ShpNumberNewFailed, ShpNameNewFailed, s)
End Property
Public Property Let ShpItemNewFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewFailed, ShpNameNewFailed.Offset(ShpNumberNewFailed, 0)), s
    ShpNumberNewFailed = ShpNumberNewFailed + 1
End Property

Public Property Get ShpItemObsolete(Optional ByVal s As String) As Boolean
    ShpItemObsolete = NameFound(ShpNumberObsolete, ShpNameObsolete, s)
End Property
Public Property Let ShpItemObsolete(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsolete, ShpNameObsolete.Offset(ShpNumberObsolete, 0)), s
    ShpNumberObsolete = ShpNumberObsolete + 1
End Property

Public Property Get ShpItemObsoleteDone(Optional ByVal s As String) As Boolean
    ShpItemObsoleteDone = NameFound(ShpNumberObsoleteDone, ShpNameObsoleteDone, s)
End Property
Public Property Let ShpItemObsoleteDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteDone, ShpNameObsoleteDone.Offset(ShpNumberObsoleteDone, 0)), s
    ShpNumberObsoleteDone = ShpNumberObsoleteDone + 1
End Property

Public Property Get ShpItemObsoleteFailed(Optional ByVal s As String) As Boolean
    ShpItemObsoleteFailed = NameFound(ShpNumberObsoleteFailed, ShpNameObsoleteFailed, s)
End Property
Public Property Let ShpItemObsoleteFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteFailed, ShpNameObsoleteFailed.Offset(ShpNumberObsoleteFailed, 0)), s
    ShpNumberObsoleteFailed = ShpNumberObsoleteFailed + 1
End Property

Public Property Get ShpItemChanged(Optional ByVal s As String) As Boolean
    ShpItemChanged = NameFound(ShpNumberChanged, ShpNameChanged, s)
End Property
Public Property Let ShpItemChanged(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChanged, ShpNameChanged.Offset(ShpNumberChanged, 0)), s
    ShpNumberChanged = ShpNumberChanged + 1
End Property

Public Property Get ShpItemChangedDone(Optional ByVal s As String) As Boolean
    ShpItemChangedDone = NameFound(ShpNumberChangedDone, ShpNameChangedDone, s)
End Property
Public Property Let ShpItemChangedDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedDone, ShpNameChangedDone.Offset(ShpNumberChangedDone, 0)), s
    ShpNumberChangedDone = ShpNumberChangedDone + 1
End Property

Public Property Get ShpItemChangedFailed(Optional ByVal s As String) As Boolean
    ShpItemChangedFailed = NameFound(ShpNumberChangedFailed, ShpNameChangedFailed, s)
End Property
Public Property Let ShpItemChangedFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedFailed, ShpNameChangedFailed.Offset(ShpNumberChangedFailed, 0)), s
    ShpNumberChangedFailed = ShpNumberChangedFailed + 1
End Property

Public Property Get ShpPrprtys(Optional ByVal s As String) As Boolean
    ShpPrprtys = NameFound(ShpNumberPrprtys, ShpNamePrprtys, s)
End Property
Public Property Let ShpPrprtys(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColPrprtys, ShpNamePrprtys.Offset(ShpNumberPrprtys, 0)), s
    ShpNumberPrprtys = ShpNumberPrprtys + 1
End Property

Public Property Get ShpPrprtysSyncExample(Optional ByVal s As String) As Boolean
    ShpPrprtysSyncExample = NameFound(ShpNumberPrprtysSyncExamples, ShpNamePrprtysSyncExamples, s)
End Property
Public Property Let ShpPrprtysSyncExample(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColPrprtysSyncExamples, ShpNamePrprtysSyncExamples.Offset(ShpNumberPrprtysSyncExamples, 0)), s
    ShpNumberPrprtysSyncExamples = ShpNumberPrprtysSyncExamples + 1
End Property


Public Property Get VbcItemAll(Optional ByVal s As String) As Boolean
    VbcItemAll = NameFound(VbcNumberAll, Intersect(ItemsColAll, VbcNumberRow), s)
End Property
Public Property Let VbcItemAll(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColAll, VbcNameTopRow.Offset(VbcNumberAll, 0)), s
    VbcNumberAll = VbcNumberAll + 1
End Property


Public Property Get VbcItemChanged(Optional ByVal s As String) As Boolean
    VbcItemChanged = NameFound(VbcNumberChanged, VbcNameChanged, s)
End Property
Public Property Let VbcItemChanged(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChanged, VbcNameChanged.Offset(VbcNumberChanged, 0)), s
    VbcNumberChanged = VbcNumberChanged + 1
End Property

Public Property Get VbcItemChangedDone(Optional ByVal s As String) As Boolean
    VbcItemChangedDone = NameFound(VbcNumberChangedDone, VbcNameChangedDone, s)
End Property
Public Property Let VbcItemChangedDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedDone, VbcNameChangedDone.Offset(VbcNumberChangedDone, 0)), s
    VbcNumberChangedDone = VbcNumberChangedDone + 1
End Property

Public Property Get VbcItemChangedFailed(Optional ByVal s As String) As Boolean
    VbcItemChangedFailed = NameFound(VbcNumberChangedFailed, VbcNameChangedFailed, s)
End Property
Public Property Let VbcItemChangedFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedFailed, VbcNameChangedFailed.Offset(VbcNumberChangedFailed, 0)), s
    VbcNumberChangedFailed = VbcNumberChangedFailed + 1
End Property

Public Property Get VbcItemNew(Optional ByVal s As String) As Boolean
    VbcItemNew = NameFound(VbcNumberNew, Intersect(ItemsColNew, VbcNumberRow), s)
End Property
Public Property Let VbcItemNew(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNew, VbcNameTopRow.Offset(VbcNumberNew, 0)), s
    VbcNumberNew = VbcNumberNew + 1
End Property

Public Property Get VbcItemNewDone(Optional ByVal s As String) As Boolean
    VbcItemNewDone = NameFound(VbcNumberNewDone, Intersect(ItemsColNewDone, VbcNumberRow), s)
End Property
Public Property Let VbcItemNewDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewDone, VbcNameNewDone.Offset(VbcNumberNewDone, 0)), s
    VbcNumberNewDone = VbcNumberNewDone + 1
End Property

Public Property Get VbcItemNewFailed(Optional ByVal s As String) As Boolean
    VbcItemNewFailed = NameFound(VbcNumberNewFailed, Intersect(ItemsColNewFailed, VbcNumberRow), s)
End Property
Public Property Let VbcItemNewFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewFailed, VbcNameNewFailed.Offset(VbcNumberNewFailed, 0)), s
    VbcNumberNewFailed = VbcNumberNewFailed + 1
End Property

Public Property Get VbcItemObsolete(Optional ByVal s As String) As Boolean
    VbcItemObsolete = NameFound(VbcNumberObsolete, Intersect(ItemsColObsolete, VbcNumberRow), s)
End Property
Public Property Let VbcItemObsolete(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsolete, VbcNameObsolete.Offset(VbcNumberObsolete, 0)), s
    VbcNumberObsolete = VbcNumberObsolete + 1
End Property

Public Property Get VbcItemObsoleteDone(Optional ByVal s As String) As Boolean
    VbcItemObsoleteDone = NameFound(VbcNumberObsoleteDone, Intersect(ItemsColObsoleteDone, VbcNumberRow), s)
End Property
Public Property Let VbcItemObsoleteDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteDone, VbcNameObsoleteDone.Offset(VbcNumberObsoleteDone, 0)), s
    VbcNumberObsoleteDone = VbcNumberObsoleteDone + 1
End Property

Public Property Get VbcItemObsoleteFailed(Optional ByVal s As String) As Boolean
    VbcItemObsoleteFailed = NameFound(VbcNumberObsoleteFailed, Intersect(ItemsColObsoleteFailed, VbcNumberRow), s)
End Property
Public Property Let VbcItemObsoleteFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteFailed, VbcNameObsoleteFailed.Offset(VbcNumberObsoleteFailed, 0)), s
    VbcNumberObsoleteFailed = VbcNumberObsoleteFailed + 1
End Property

Public Property Get RngItemAll(Optional ByVal s As String) As Boolean
    RngItemAll = NameFound(RngNumberAll, Intersect(ItemsColAll, RngNumberRow), s)
End Property
Public Property Let RngItemAll(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColAll, RngNameTopRow.Offset(RngNumberAll, 0)), s
    RngNumberAll = RngNumberAll + 1
End Property

Public Property Get RngItemNew(Optional ByVal s As String) As Boolean
    RngItemNew = NameFound(RngNumberNew, Intersect(ItemsColNew, RngNumberRow), s)
End Property
Public Property Let RngItemNew(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNew, RngNameTopRow.Offset(RngNumberNew, 0)), s
    RngNumberNew = RngNumberNew + 1
End Property

Public Property Get RngItemNewDone(Optional ByVal s As String) As Boolean
    RngItemNewDone = NameFound(RngNumberNewDone, Intersect(ItemsColNew, RngNumberRow), s)
End Property
Public Property Let RngItemNewDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewDone, RngNameNewDone.Offset(RngNumberNewDone, 0)), s
    RngNumberNewDone = RngNumberNewDone + 1
End Property

Public Property Get RngItemNewFailed(Optional ByVal s As String) As Boolean
    RngItemNewFailed = NameFound(RngNumberNewFailed, Intersect(ItemsColNewFailed, RngNumberRow), s)
End Property
Public Property Let RngItemNewFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewFailed, RngNameNewFailed.Offset(RngNumberNewFailed, 0)), s
    RngNumberNewFailed = RngNumberNewFailed + 1
End Property

Public Property Get RngItemObsolete(Optional ByVal s As String) As Boolean
    RngItemObsolete = NameFound(RngNumberObsolete, Intersect(ItemsColObsolete, RngNumberRow), s)
End Property
Public Property Let RngItemObsolete(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsolete, RngNameObsolete.Offset(RngNumberObsolete, 0)), s
    RngNumberObsolete = RngNumberObsolete + 1
End Property

Public Property Get RngItemObsoleteDone(Optional ByVal s As String) As Boolean
    RngItemObsoleteDone = NameFound(RngNumberObsoleteDone, Intersect(ItemsColObsoleteDone, RngNumberRow), s)
End Property
Public Property Let RngItemObsoleteDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteDone, RngNameObsoleteDone.Offset(RngNumberObsoleteDone, 0)), s
    RngNumberObsoleteDone = RngNumberObsoleteDone + 1
End Property

Public Property Get RngItemObsoleteFailed(Optional ByVal s As String) As Boolean
    RngItemObsoleteFailed = NameFound(RngNumberObsoleteFailed, Intersect(ItemsColObsoleteFailed, RngNumberRow), s)
End Property
Public Property Let RngItemObsoleteFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteFailed, RngNameObsoleteFailed.Offset(RngNumberObsoleteFailed, 0)), s
    RngNumberObsoleteFailed = RngNumberObsoleteFailed + 1
End Property

Public Property Get RngItemChanged(Optional ByVal s As String) As Boolean
    RngItemChanged = NameFound(RngNumberChanged, Intersect(ItemsColChanged, RngNumberRow), s)
End Property
Public Property Let RngItemChanged(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChanged, RngNameChanged.Offset(RngNumberChanged, 0)), s
    RngNumberChanged = RngNumberChanged + 1
End Property

Public Property Get RngItemChangedDone(Optional ByVal s As String) As Boolean
    RngItemChangedDone = NameFound(RngNumberChangedDone, Intersect(ItemsColChangedDone, RngNumberRow), s)
End Property
Public Property Let RngItemChangedDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedDone, RngNameChangedDone.Offset(RngNumberChangedDone, 0)), s
    RngNumberChangedDone = RngNumberChangedDone + 1
End Property

Public Property Get RngItemChangedFailed(Optional ByVal s As String) As Boolean
    RngItemChangedFailed = NameFound(RngNumberChangedFailed, Intersect(ItemsColChangedFailed, RngNumberRow), s)
End Property
Public Property Let RngItemChangedFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedFailed, RngNameChangedFailed.Offset(RngNumberChangedFailed, 0)), s
    RngNumberChangedFailed = RngNumberChangedFailed + 1
End Property


























Public Property Let WshItemAll(ByVal s As String)
    AddItem Intersect(ItemsColAll, WshNameTopRow.Offset(WshNumberAll, 0)), s
    WshNumberAll = WshNumberAll + 1
End Property

Public Property Get WshItemNew(Optional ByVal s As String) As Boolean
    WshItemNew = NameFound(WshNumberNew, Intersect(ItemsColNew, WshNumberRow), s)
End Property
Public Property Let WshItemNew(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNew, WshNameTopRow.Offset(WshNumberNew, 0)), s
    WshNumberNew = WshNumberNew + 1
End Property

Public Property Get WshItemNewDone(Optional ByVal s As String) As Boolean
    WshItemNewDone = NameFound(WshNumberNewDone, Intersect(ItemsColNew, WshNumberRow), s)
End Property
Public Property Let WshItemNewDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewDone, WshNameNewDone.Offset(WshNumberNewDone, 0)), s
    WshNumberNewDone = WshNumberNewDone + 1
End Property

Public Property Get WshItemNewFailed(Optional ByVal s As String) As Boolean
    WshItemNewFailed = NameFound(WshNumberNewFailed, Intersect(ItemsColNewFailed, WshNumberRow), s)
End Property
Public Property Let WshItemNewFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColNewFailed, WshNameNewFailed.Offset(WshNumberNewFailed, 0)), s
    WshNumberNewFailed = WshNumberNewFailed + 1
End Property

Public Property Get WshItemObsolete(Optional ByVal s As String) As Boolean
    WshItemObsolete = NameFound(WshNumberObsolete, Intersect(ItemsColObsolete, WshNumberRow), s)
End Property
Public Property Let WshItemObsolete(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsolete, WshNameObsolete.Offset(WshNumberObsolete, 0)), s
    WshNumberObsolete = WshNumberObsolete + 1
End Property

Public Property Get WshItemObsoleteDone(Optional ByVal s As String) As Boolean
    WshItemObsoleteDone = NameFound(WshNumberObsoleteDone, Intersect(ItemsColObsoleteDone, WshNumberRow), s)
End Property
Public Property Let WshItemObsoleteDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteDone, WshNameObsoleteDone.Offset(WshNumberObsoleteDone, 0)), s
    WshNumberObsoleteDone = WshNumberObsoleteDone + 1
End Property

Public Property Get WshItemObsoleteFailed(Optional ByVal s As String) As Boolean
    WshItemObsoleteFailed = NameFound(WshNumberObsoleteFailed, Intersect(ItemsColObsoleteFailed, WshNumberRow), s)
End Property
Public Property Let WshItemObsoleteFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColObsoleteFailed, WshNameObsoleteFailed.Offset(WshNumberObsoleteFailed, 0)), s
    WshNumberObsoleteFailed = WshNumberObsoleteFailed + 1
End Property

Public Property Get WshItemChanged(Optional ByVal s As String) As Boolean
    WshItemChanged = NameFound(WshNumberChanged, Intersect(ItemsColChanged, WshNumberRow), s)
End Property
Public Property Let WshItemChanged(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChanged, WshNameChanged.Offset(WshNumberChanged, 0)), s
    WshNumberChanged = WshNumberChanged + 1
End Property

Public Property Get WshItemChangedDone(Optional ByVal s As String) As Boolean
    WshItemChangedDone = NameFound(WshNumberChangedDone, Intersect(ItemsColChangedDone, WshNumberRow), s)
End Property
Public Property Let WshItemChangedDone(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedDone, WshNameChangedDone.Offset(WshNumberChangedDone, 0)), s
    WshNumberChangedDone = WshNumberChangedDone + 1
End Property

Public Property Get WshItemChangedFailed(Optional ByVal s As String) As Boolean
    WshItemChangedFailed = NameFound(WshNumberChangedFailed, Intersect(ItemsColChangedFailed, WshNumberRow), s)
End Property
Public Property Let WshItemChangedFailed(Optional ByVal s As String, ByVal b As Boolean)
    AddItem Intersect(ItemsColChangedFailed, WshNameChangedFailed.Offset(WshNumberChangedFailed, 0)), s
    WshNumberChangedFailed = WshNumberChangedFailed + 1
End Property










Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "wsSync" & "." & sProc
End Function

Private Function NameFound(ByVal l As Long, _
                           ByVal r As Range, _
                           ByVal s As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when the string (s) is found within 1 to (l) in the cells of
' range (r).
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To l - 1
        If r.Offset(i, 0).Value = s Then
            NameFound = True
            Exit For
        End If
    Next i

End Function

Private Function CBoolString(ByVal b As Boolean) As String
    If b _
    Then CBoolString = "'True" _
    Else CBoolString = "'False"
End Function

Private Function CStringBool(ByVal s As String) As Boolean
    If s = "True" _
    Then CStringBool = True _
    Else CStringBool = False
End Function

Public Sub Clear()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Clear"
    
    On Error GoTo eh
    With Me
        .NmeSyncDone = False
        .RefSyncDone = False
        .RngSyncDone = False
        .ShpSyncDone = False
        .VbcSyncDone = False
        .WshSyncDone = False
    End With
    
    ClearKind ShpNumberRow, ShpNameTopRow
    ClearKind RefNumberRow, RefNameTopRow
    ClearKind WshNumberRow, WshNameTopRow
    ClearKind VbcNumberRow, VbcNameTopRow
    ClearKind RngNumberRow, RngNameTopRow
    ClearKind NmeNumberRow, NmeNameTopRow

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ClearKind(ByVal number_row As Range, _
                     ByVal name_top_row As Range)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "ClearKind"
    
    On Error GoTo eh
    
    Dim cel As Range
    Dim i   As Long
    Dim j   As Long
    
    For Each cel In Intersect(number_row, NumbersAndNamesCols).Cells
        j = Max(j, cel.Value)
    Next cel
    
    If j > 0 Then
        Me.Activate
        Me.Range(name_top_row, name_top_row.Offset(j - 1)).EntireRow.Delete
    End If
    
    For Each cel In Intersect(number_row, NumbersAndNamesCols).Cells
        With cel
            If Not .Locked Then .Value = 0
        End With
    Next cel

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub AddItem(ByVal r As Range, _
                    ByVal v As Variant)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "AddItem"
    
    On Error GoTo eh
    Dim rBottom As Range
    Dim rName   As Range
    
    Set rBottom = Intersect(r.EntireRow, KindOfSynchedItems.EntireColumn)
    With rBottom
        If .Value <> vbNullString Then
            .Rows(1).EntireRow.Insert
            Set rName = Intersect(.Offset(-1, 0).EntireRow, r.EntireColumn)
        Else
            Set rName = Intersect(.EntireRow, r.EntireColumn)
        End If
    End With
    
    With rName
        With .Font
            .Color = rgbBlack
            .Size = 8
            .Name = "Courier New"
        End With
        .HorizontalAlignment = xlLeft
        .Value = v
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

