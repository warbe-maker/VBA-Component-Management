Attribute VB_Name = "mCompManTest"
Option Explicit
' ------------------------------------------------------------------------------
' Standard module mTest: Provides the means to test the services of the CompMan
' ====================== AddIn Workbook.
'
' W. Rauschenberger Berlin, Jan 2021
' ------------------------------------------------------------------------------
Private cTest   As clsTestService

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
    
    mSyncTest.TestSync cTest.Regression
    
xt: mBasic.EoP ErrSrc(PROC)
    mTrc.Dsply
    Exit Sub
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
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
    
eh: Select Case mMe.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

