VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ---------------------------------------------------------------------------
' Common Component Workbook CompManDev.xlsm
' - When setup this Workbook also exists as AddIn Workbook CompMan.xlma
' - When renewed, this Workbook is open as Addin Workbook CompMan.xlma
'
' Properties:
' - AddInInstanceFullName Full path/file name of the Addin Instance
' - AddInInstanceName     File name of the Addin instance
' - DevlpInstanceName     File name of the Development and Test
' -                       instance
' - IsAddInInstance       Is TRUE when ThisWorkbook is the
'                         AddIn instance
' - IsDevlpInstance       Is TRUE when ThisWorkbook is the
'                         development and test instance
' - xlAddInFormat         The file format for the SaveAs operation
'
' W. Rauschenberger, Berlin Dec 2020
' See Github repo
' ---------------------------------------------------------------------------

Public Property Get HOSTED_RAWS() As String:            HOSTED_RAWS = "":                                                End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "wbCompMan" & "." & sProc
End Function

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If mMe.IsDevInstnc Then mMe.ControlItemRenewRemove
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Const PROC      As String = "Workbook_BeforeSave"
    
    On Error GoTo eh
    Dim wb As Workbook

    If Not mMe.IsDevInstnc Then Exit Sub
    
    Set wb = ThisWorkbook
    mErH.BoP ErrSrc(PROC)
        
    Application.EnableEvents = False
    wb.Save
    Application.EnableEvents = True
    mCompMan.ExportChangedComponents wb, HOSTED_RAWS ' includes management of updated used Common Components
    If Not wb.Saved Then
        Application.EnableEvents = False
        wb.Save        ' usually not required
        Application.EnableEvents = True
    End If
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Private Sub Workbook_Open()
    Const PROC = "Workbook_Open"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim cLog As New clsLog: cLog.Reset: Set cLog = Nothing
    
    mErH.BoP ErrSrc(PROC)
    
    If mMe.IsDevInstnc Then
        mMe.ControlItemRenewAdd
        mMe.ConfigAsserted ' assert a valid CompMan basic configuratuion
        If mMe.AddInInstncWrkbkIsOpen Then
            '~~ Essential! When the Development instance Workbook is opened, used Common Components
            '~~ the origin code has changed can only be updated when the Addin instance Workbook is
            '~~ open as well. However, since the VBProject has no reference to the Addin the Addin
            '~~ will not be opened when the Development instance is the only Workbook open.
            '~~ I.e. only when a Workbook referring to the Addin is opened the Addin will be opened as well.
            mMe.UpdateUsedCommComps
        End If
    End If
        
    With ThisWorkbook
        '~~ Make sure the VBProject name is the Workbook's base name
        '~~ Note that this is required to correctly identify (save, remove, restore) any reference
        '~~ to the Addin in open Workbook when Renew or SaveAsAddin is performed (see module mAddin)
        If .VBProject.name <> fso.GetBaseName(.name) Then .VBProject.name = fso.GetBaseName(.name)
        '~~ Unporotecting (the Addin) allows directly modify code in it - and saving it subsequently
        '~~ as new Development instance (see method SaveasDev in module mAddin)
        .Unprotect
    End With
    
    
xt: mErH.BoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

