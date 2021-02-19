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
' ------------------------------------------------------------------
' !! This Workbook exists as two instances:                       !!
' !! 1. The development instance CompManDev.xlsb                  !!
' !! 2. The Addin instance CompMan.xlma which is setup/renewed    !!
' !!    via the development instance service mMe.RenewAddIn       !!
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
' ------------------------------------------------------------------

Friend Property Get HOSTED_RAWS() As String:    HOSTED_RAWS = "":   End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "wbAddIn" & "." & sProc
End Function

Private Sub Workbook_Open()
' -------------------------------------------------------------------------
' Attention: ThisWorkbook's raw clones (i.e. the used common components)
'            are only updated when the CompMan Addin is active. This is
'            only the case when at least one Workbook referring to it is
'            open (this development instance cannot refer to 'its' Addin
'            instance because it may not exist - in case of a 'cold start'.
'            On the other hand, it will be essential that the Addin is
'            established with up-to-date used raw clones.
'            Conclusion: For a 'cold start', i.e. the development instance
'            is opened the very first time it is assumed up-to-date. Later
'            on, when a used clone's raw is changed - that is when the
'            Addin will be active - the development instance will be opened
'            to get its cloned raws updated and the Addin instance will be
'            renewed.
' -------------------------------------------------------------------------
    Const PROC = "Workbook_Open"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    mErH.BoP ErrSrc(PROC)
    
    If mMe.IsDevInstnc Then
        '~~ Assert a valid basic CompMan configuration (confirmation dialog only when invalid)
        '~~ When terminated with an invalid configuration services will be denied
        mMe.BasicConfig
        If mMe.AddInInstncWrkbkIsOpen Then
            '~~ The update of any Clone-Components the Raw-Component had changed requires the CompMan-Addin.
            '~~ When the development instance is the only Workbook open the CompMan-Addin might not be available
            '~~ and thus no Clone-Components will be updated.
            Application.StatusBar = vbNullString
            mMe.UpdateRawClones
        End If
    End If
        
    With ThisWorkbook
        '~~ Make sure the VBProject name is the Workbook's base name
        '~~ Note that this is required to correctly identify (save, remove, restore) any reference
        '~~ to the Addin in open Workbook when RenewByImport or SaveAsAddin is performed (see module mAddin)
        If .VbProject.name <> fso.GetBaseName(.name) Then .VbProject.name = fso.GetBaseName(.name)
        '~~ Unporotecting (the Addin) allows directly modify code in it - and saving it subsequently
        '~~ as new Development instance (see method SaveasDev in module mAddin)
        .Unprotect
    End With
    
    
xt: mErH.BoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Const PROC = "Workbook_BeforeSave"
    
    On Error GoTo eh
    Dim wb As Workbook

    If Not mMe.IsDevInstnc Then Exit Sub
    
    Set wb = ThisWorkbook
    mErH.BoP ErrSrc(PROC)
        
    Application.EnableEvents = False
    Application.StatusBar = "Saving Workbook '" & wb.name & "' please wait, this may take a few seconds"
    wb.Save
    Application.StatusBar = vbNullString
    mCompMan.ExportChangedComponents wb, HOSTED_RAWS ' includes management of updated used Common Components
    If Not wb.Saved Then
        wb.Save        ' usually not required
    End If
    
xt: Application.EnableEvents = True
    mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

