VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fSelect 
   Caption         =   "Select the Components for being exported/backed up"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5190
   OleObjectBlob   =   "fSelect.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "fSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dct                 As Dictionary
Private wbHost              As Workbook
Private bWhenChangedOnly    As Boolean

Public Property Let Host(ByVal wb As Workbook):     Set wbHost = wb:                    End Property
Public Property Get SelectedComps() As Dictionary:  Set SelectedComps = dct:            End Property
Public Property Get WhenChangedOnly() As Boolean:   WhenChangedOnly = bWhenChangedOnly: End Property

Private Sub cbxWhenChangedOnly_Click()
    With Me
        bWhenChangedOnly = .cbxWhenChangedOnly
    End With
End Sub

Private Sub cmbOk_Click()
Dim i As Long

    With Me.lbSelect
        For i = 0 To .ListCount - 1
            If .Selected(i) = False Then
                dct.Remove .List(i)
            End If
        Next i
    End With
    Me.Hide
    
End Sub

Private Sub UserForm_Activate()

    Dim vbc     As VBComponent
    Dim v       As Variant

    If dct Is Nothing Then Set dct = New Dictionary Else dct.RemoveAll
    
    For Each vbc In wbHost.VBProject.VBComponents
        mDct.DctAdd add_dct:=dct, add_key:=vbc.Name, add_item:=vbc, add_seq:=seq_ascending
    Next vbc
    
    With Me
        For Each v In dct
            .lbSelect.AddItem v
        Next v
        .laCompsInWorkbook.caption = Replace(.laCompsInWorkbook.caption, "<wb>", wbHost.Name)
    End With

End Sub

