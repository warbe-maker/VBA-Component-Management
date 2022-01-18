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
' ------------------------------------------------------------------------------
' !! This Workbook exists as two instances:                                   !!
' !! 1. The development instance CompManDev.xlsb                              !!
' !! 2. The Addin instance CompMan.xlma which is setup/renewed via the        !!
' !!    development instance service mMe.RenewAddIn                           !!
'
' See: https://github.com/warbe-maker/Excel-VB-Components-Management-Services
'
' W. Rauschenberger, Berlin Dec 2020
' ------------------------------------------------------------------------------
Private Const HOSTED_RAWS = "mCompManClient"

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "wbAddIn" & "." & sProc
End Function

Private Sub Workbook_Open()
' ------------------------------------------------------------------------------
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
' ------------------------------------------------------------------------------
    Const PROC = "Workbook_Open"
    
    On Error GoTo eh
       
    If mMe.IsDevInstnc Then
        mMe.BasicConfig bc_silent:=True
        '~~ ------------------------------------------------------------------
        '~~ CompMan Workbook_Open service 'UpdateOutdatedCommonComponents':
        '~~ Executed by the development instance when open or by the Addin
        '~~ when established and automatically available when referenced by
        '~~ the VB-Project
        mCompManClient.CompManService "UpdateOutdatedCommonComponents", HOSTED_RAWS
        '~~ ------------------------------------------------------------------
        mMe.DisplayStatus
        Application.ScreenUpdating = False
        With ActiveWindow
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
        End With
        wsAddIn.Activate
        Application.DisplayFormulaBar = False
        ActiveWindow.DisplayHeadings = False
        wsAddIn.Range("rngDisplayed").Select
        ActiveWindow.Zoom = True
        wsAddIn.Range("A1").Select
        wsAddIn.Range("celSelect").Select
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Const PROC = "Workbook_BeforeSave"
    
    On Error GoTo eh
    '~~ ------------------------------------------------------------------
    '~~ 'ExportChangedComponents' service, preferrably performed by the
    '~~ CompMan Addin *), or when not available by this Workbook itself.
    '~~ *) automatically available only when a Workbook is open which has
    '~~    in the VBProject a Reference to the CompMan Addin.
    If mMe.IsDevInstnc Then
        mMe.BasicConfig bc_silent:=True
           mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
        If Not ThisWorkbook.Saved Then
            Application.EnableEvents = False
            ThisWorkbook.Save
            Application.EnableEvents = True
        End If
    End If
    '~~ ------------------------------------------------------------------
    
xt: Application.EnableEvents = True
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

