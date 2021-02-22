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
Private Const HOSTED_RAWS = ""
Private Const COMPMAN_BY_ADDIN  As String = "CompMan.xlam!mCompMan."
Private Const COMPMAN_BY_DEVLP  As String = "CompMan.xlsb!mCompMan."

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
        '~~ ------------------------------------------------------------------
        '~~ CompMan Workbook_Open service 'UpdateRawClones':
        '~~ Executed by the Addin *) or via the development instance when open
        '~~ *) automatically available only when referenced by the VB-Project
        CompManService "UpdateRawClones"
        '~~ ------------------------------------------------------------------
    End If
        
    ThisWorkbook.Unprotect
    
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
    Application.StatusBar = "Saving Workbook '" & ThisWorkbook.name & "' please wait, this may take a few seconds"
    Application.EnableEvents = False
    ThisWorkbook.Save
    Application.StatusBar = vbNullString
    '~~ ------------------------------------------------------------------
    '~~ CompMan Workbook_Before_Save service 'ExportChangedComponents':
    '~~ Executed by the Addin *) or via the development instance when open
    '~~ *) automatically available only when referenced by the VB-Project
    CompManService "ExportChangedComponents"
    '~~ ------------------------------------------------------------------
    If Not ThisWorkbook.Saved Then ThisWorkbook.Save
    
xt: Application.EnableEvents = True
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Private Function CompManService(ByVal service As String) As Boolean
' ----------------------------------------------------------------------------
' Execution of the CompMan service (service) preferrably via the CompMan-Addin
' or when not available alternatively via the CompMan's development instance.
' ----------------------------------------------------------------------------
    On Error Resume Next
    Application.Run COMPMAN_BY_ADDIN & service, ThisWorkbook, HOSTED_RAWS
    If Err.Number = 1004 Then
        Application.StatusBar = "CompMan-Service: '" & service & "' not available by '" & COMPMAN_BY_ADDIN & "'!"
    End If
End Function

