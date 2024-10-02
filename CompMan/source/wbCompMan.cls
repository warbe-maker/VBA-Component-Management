VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbCompMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ------------------------------------------------------------------------------
' || This Workbook may exist as two instances:
' || 1. CompMan.xlsb as the development instance
' || 2. CompMan.xlma as Add-in instance  when one has been setup via the
' ||    development instance service mMe.Renew___AddIn
'
' See: https://github.com/warbe-maker/Excel-VB-Components-Management-Services
'
' W. Rauschenberger, Berlin Dec 2020
' ------------------------------------------------------------------------------
Private Const HOSTED_RAWS        As String = "mCompManClient,mCodingGuidelines,mCodingRules"
Private Const PUBLIC_PROC_COPIES As String = "clsTestAid.Arr:mBasic, clsTestAid.FileAsArray:mVarTrans"

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = Me.CodeName & "." & sProc
End Function

Private Sub Workbook_BeforeClose(Cancel As Boolean)
' ------------------------------------------------------------------------------
' Ensure the CompMan Add-in is not (no longer) paused when the Development
' Instance Workbook is closed.
' ------------------------------------------------------------------------------
   
    If mMe.IsDevInstnc Then
        If mAddin.IsOpen Then
            mAddin.Paused = False
        End If
    End If
    If Application.Workbooks.Count = 1 Then
        mCompManMenuVBE.MenuRemove
    End If
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
' ------------------------------------------------------------------------------
' Prevents a Workbook with an invalid configuration is saved.
' ------------------------------------------------------------------------------
   
    If mMe.AssertedServicingEnabled Then
        If Not wsConfig.Verified Then
            wsConfig.Activate
            Cancel = True
        End If
    End If
    Application.CutCopyMode = False
    Application.StatusBar = "Workbook is being saved"
    
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, HOSTED_RAWS, PUBLIC_PROC_COPIES
End Sub

Private Sub Workbook_Open()
' ----------------------------------------------------------------------------
' Attention:  ThisWorkbook's raw clones (i.e. the used common components) are
'             only updated when the CompMan Add-in is active.
' Conclusion: For a 'cold start' of CompMan, i.e. the development instance is
'             opened the very first time it is assumed up-to-date. Later on,
'             when a used clone's raw is changed - that is when the Add-in will
'             be active - the development instance will be opened to get its
'             cloned raws updated and the Add-in instance will be renewed.
' ------------------------------------------------------------------------------
    Const PROC = "Workbook_Open"
       
    On Error GoTo eh
    If Not mMe.AssertedServicingEnabled Then GoTo xt
    
    If mMe.IsDevInstnc Then
        '~~ ------------------------------------------------------------------
        '~~ CompMan Workbook_Open service 'UpdateOutdatedCommonComponents':
        '~~ Executed by the development instance when open or by the Add-in
        '~~ when established and automatically available when referenced by
        '~~ the VB-Project
        mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS, PUBLIC_PROC_COPIES
        '~~ ------------------------------------------------------------------
        Application.Wait Now() + 0.0000001
        '~~ ------------------------------------------------------------------
        '~~ When the Workbook is opened from within the configured 'Serviced
        '~~ Sync Target' folder it will be synchronized with its corresponding
        '~~ raw (the Sync-Source-Workbook) located in the 'Serviced Dev and
        '~~ Test' folder. When the required preconditions are not met the
        '~~ service will be denied.
'        mCompManClient.CompManService mCompManClient.SRVC_SYNCHRONIZE
        '~~ ------------------------------------------------------------------
    End If
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


