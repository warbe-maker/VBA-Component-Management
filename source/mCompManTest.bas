Attribute VB_Name = "mCompManTest"
Option Explicit
' ------------------------------------------------------------------------------
' Standard module mTest: Provides the means to test the services of the CompMan
' ====================== AddIn Workbook.
'
' W. Rauschenberger Berlin, Jan 2021
' ------------------------------------------------------------------------------
Private Const TEST_CHANGE = "' Test code change"

Private cTest   As clsTestService
Private wbTest  As Workbook
Private wbSrc   As Workbook
Private wbTrgt  As Workbook
Private vbc     As VBComponent
Private vbcm    As CodeModule
Private fso     As New FileSystemObject

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompManTest" & "." & sProc
End Function

Private Sub Regression()
' -----------------------------------------------------------------
' Test of all Public CompMan services.
' Note: Test of "UpdateOutdatedCommonComponents" is still pending
'       since it requires the setup of a Test-Workbook very similar
'       to those setup for the "SynchronizeVBProjects" test
' -----------------------------------------------------------------
    Const PROC = "Regression"
    
    On Error GoTo eh
    Set cTest = New clsTestService
    cTest.Regression = True
    Set Trc = New clsTrc ' if not within regression-test
    mBasic.BoP ErrSrc(PROC)
    
    mCompManTest.Test_ExportChanged
    mSyncTest.TestSync cTest.Regression
    
xt: mBasic.EoP ErrSrc(PROC)
    mTrc.Dsply
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_ExportChanged()
    Const PROC = "Test_ExportChanged"
    
    On Error GoTo eh
    Dim Comp        As New clsComp
    Dim wbActive    As Workbook
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' if not within regression-test
    mBasic.BoP ErrSrc(PROC)
   
    Services.Initiate ErrSrc(PROC), ThisWorkbook
    LogServiced.Title "Export Changed Components Test"
    mCompMan.ExportChangedComponents ThisWorkbook, "mCompManClient"

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_LogServiced()
    Const PROC = "Test_LogServiced"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    
    Services.Initiate PROC, ThisWorkbook
    With LogServiced
        .Title ErrSrc(PROC)
        Services.ServicedItem = " <component-name> "
        Services.LogServicedEntry "Tested"
        mMsg.Box Title:="Test-LogServiced:" _
               , Prompt:=mFso.FileTxt(.LogFile.Path) _
               , box_monospaced:=True
        If fso.FileExists(.LogFile.Path) Then fso.DeleteFile .LogFile.Path
    End With
    
xt: Set Services = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_RenewByImport(ByVal rnc_exp_file_full_name, _
                               ByVal rnc_vbc_name As String)
' ------------------------------------------------------------------------------
' This test procedure is exclusively initiated within the
' 'CompMan-Development-Instance-Workbook and executed by' the 'CompMan Add-in'
' which needs to be open. When these conditions are not met a message is
' displayed.
' ------------------------------------------------------------------------------
    Const PROC = "Test_RenewByImport"
    
    On Error GoTo eh
    Dim Comp        As New clsComp
    Dim wbActive    As Workbook
    Dim wbTemp      As Workbook
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' if not within regression-test
    mBasic.BoP ErrSrc(PROC)
    If mMe.IsDevInstnc Then GoTo xt
    
    Services.Initiate ErrSrc(PROC), ThisWorkbook
    LogServiced.Title "Test RenewByImport"
    
    With Comp
        .CompName = rnc_vbc_name
        Services.ServicedItem = .VBComp
        
        If .Wrkbk Is ActiveWorkbook Then
            Set wbActive = ActiveWorkbook
            Set wbTemp = Workbooks.Add ' Activates a temporary Workbook
            LogServiced.Entry "Active Workbook de-activated by creating a temporary Workbook"
        End If
            
        mUpdate.ByReImport b_wbk_target:=.Wrkbk _
                         , b_vbc_name:=.CompName _
                         , b_exp_file:=rnc_exp_file_full_name
    End With
    
xt: If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
        LogServiced.Entry "Temporary created Workbook closed without save"
        Set wbTemp = Nothing
        If Not ActiveWorkbook Is wbActive Then
            wbActive.Activate
            LogServiced.Entry "De-activated Workbook '" & wbActive.Name & "' re-activated"
            Set wbActive = Nothing
        Else
            LogServiced.Entry "Workbook '" & wbActive.Name & "' re-activated by closing the temporary created Workbook"
        End If
    End If
    Set Comp = Nothing
    Set Services = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_UpdateOutdatedCommonComponents()
    Const PROC  As String = "Test_UpdateOutdatedCommonComponents"
    
    On Error GoTo eh
    Dim AddinService    As String
    Dim AddInStatus     As String
    
    If Trc Is Nothing Then Set Trc = New clsTrc ' if not within regression-test
    mBasic.BoP ErrSrc(PROC)
    If Trc Is Nothing Then Set Trc = New clsTrc ' if not within regression-test
    If Services.Denied(mCompManClient.SRVC_UPDATE_OUTDATED) Then GoTo xt

    AddinService = mAddin.WbkName & "!mCompMan.UpdateOutdatedCommonComponents"
    If mAddin.IsOpen Then
        AddInStatus = " (currently the case) "
    Else
        AddInStatus = " (currently  " & mBasic.Spaced("not") & "  the case) "
    End If
    
    If mMe.IsDevInstnc Then
        mBasic.BoP ErrSrc(PROC)
        
        On Error Resume Next
        Application.Run AddinService _
                      , ThisWorkbook
        
        If Err.Number = 1004 Then
            MsgBox Title:="CompMan Add-in not open (required for test: " & PROC & "!" _
                 , Prompt:="Application.Run " & vbLf & vbLf & AddinService & vbLf & vbLf & "failed because the 'CompMan Add-in' is not open!" _
                 , Buttons:=vbExclamation
        End If
        mBasic.EoP ErrSrc(PROC)
    Else
        MsgBox Title:="Test " & PROC & " not executed!" _
             , Prompt:="Executions of this test must not be performed 'within' the 'CompMan Add-in' Workbook." & vbLf & vbLf & _
                       "The test requires the 'CompMan Add-in' (" & mAddin.WbkName & ") is open " & AddInStatus & " but must be performed " & _
                       "from within the development instance (" & mMe.DevInstncFullName & ")." _
             , Buttons:=vbExclamation
    End If
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_Template_Proc()

    Const PROC = "Test_RevNoIncrease"
        
    If Trc Is Nothing Then ' if not executed within a regression-test
        Set Trc = New clsTrc
        Trc.NewFile ' temp trace-log-file, with default name for this test only
    End If
    mBasic.BoP ErrSrc(PROC)
    
xt: mBasic.EoP ErrSrc(PROC)
    '~~ show trace and delete trace-log-file
    Trc.Dsply
    Stop ' continue when display is closed
    fso.DeleteFile Trc.FileFullName
    Set Trc = Nothing
    
End Sub
