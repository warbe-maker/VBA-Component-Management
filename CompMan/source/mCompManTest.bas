Attribute VB_Name = "mCompManTest"
Option Explicit
' ------------------------------------------------------------------------------
' Standard module mTest: Provides the means to test the services of the CompMan
' ====================== AddIn Workbook.
'
' W. Rauschenberger Berlin, Jan 2021
' ------------------------------------------------------------------------------
Private TestAid     As New clsTestAid
Private aTestComps  As Variant
Private wbkServiced As Workbook
Private sTestComp   As String
Private sTestFolder As String

Private Function CompSelected(ByVal c_msg As String, _
                     Optional ByRef c_comp As String) As String
    c_comp = InputBox(c_msg, "Test Component precondition")
    CompSelected = c_comp
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCompManTest" & "." & sProc
End Function

Private Property Let ServicedWrkbk(ByVal s_wbk As Workbook):    Set wbkServiced = s_wbk:            End Property

Private Property Get ServicedWrkbk() As Workbook:               Set ServicedWrkbk = wbkServiced:    End Property

Private Sub Prepare()
    
    If TestAid Is Nothing Then Set TestAid = New clsTestAid
    If Trc Is Nothing Then Set Trc = New clsTrc ' if not within regression-test
    If mErH.Regression _
    Then Trc.FileFullName = TestAid.TestFolder & "\RegressionExec.trc" _
    Else Trc.FileFullName = TestAid.TestFolder & "\Exec.trc"
    TestAid.ModeRegression = mErH.Regression

End Sub

Private Sub Regression()
' -----------------------------------------------------------------
' Test of all Public CompMan services.
' Note: Test of "UpdateOutdatedCommonComponents" is still pending
'       since it requires the setup of a Test-Workbook very similar
'       to those setup for the "SynchronizeVBProjects" test
' -----------------------------------------------------------------
    Const PROC = "Regression"
    
    On Error GoTo eh
    
    mErH.Regression = True
    Prepare
    TestAid.TestHeadLineRegression = "Test of the CompMan services update and export"
    TestAid.CleanUp "Result*"
    mBasic.BoP ErrSrc(PROC)
    
    mCompManTest.Test_0100_FirstTimeServiced
    
xt: mBasic.EoP ErrSrc(PROC)
    Trc.Dsply
    TestAid.ResultSummaryLog
    Set Trc = Nothing
    Set TestAid = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function TestWorkbook(ByVal t_folder As String) As String
    Dim fle As File
    
    For Each fle In FSo.GetFolder(t_folder).Files
        If FSo.GetExtensionName(fle.Path) Like "xl*" Then
            TestWorkbook = fle.Path
            Exit For
        End If
    Next fle
    
End Function

Public Sub Test_0100_FirstTimeServiced()
' -----------------------------------------------------------------
' Covers: - A Common Component requiring a manual interaction for
'           being registered as "used"
'         - A used Common Component initially outdated, requiring
'           a manual update confirmation
'         - A hosted Common Component, exported for the first time
'           and registered pending thereby
' -----------------------------------------------------------------
    Const PROC = "Test_0100_FirstTimeServiced"
    
    Dim Comp As New clsComp
     
    On Error GoTo eh
    Prepare
    Test_0100_2CleanUp ' cleanup any remains from a previous uncompleted test
    Test_0100_1SetUp
    
    With TestAid
        .TestId = "0100" ' basic test number for the environment setup
        .TestHeadLine = "First time serviced Workbook/VBProject"
        
        mCompMan.ServiceInitiate wbkServiced, PROC
        
        '==========================================================================
        .TestId = "0100-1"
        .TestedComp = "mHskpng"
        .TestedProc = "CommCompsServicedKindOf"
        .TestedType = "Sub"
        
        '~~ Assert precondition
        .Verification = "Precondition: 0 Common Components registered in CommComps.dat"
        .ResultExpected = 0
        .Result = CommonServiced.Components.Count
        '-------------------------------------
        .RequiredInteraction = "Confirm ""mBasic"" as  u s e d  Common Component!"
        mHskpng.CommCompsServicedKindOf
        '-------------------------------------
        .Verification = "Verification: Houskeeping resulted in ""mBasic"" registered as ""used"" Common Component"
        .ResultExpected = True
        .Result = CommonServiced.IsUsedCommComp("mBasic")
        
        '==========================================================================
        .TestId = "0100-2"
        .TestedComp = "mCommComps"
        .TestedProc = "Update"
        .TestedType = "Sub"
        
        .Verification = "Precondition: Common Component ""mBasic"" is outdated"
        .ResultExpected = True
        Comp.CompName = "mBasic"
        .Result = Not Comp.CodeCrrent.Meets(Comp.CodePublic)
        '----------------
        .RequiredInteraction = "Confirm update of the outdated Common Component ""mBasic""!"
        mCommComps.Update
        '----------------
        .Verification = "Verification: Common Component ""mBasic"" is up-to-date (code current meets code public)"
        .ResultExpected = True
        Comp.CompName = "mBasic"
        .Result = Comp.CodeCrrent.Meets(Comp.CodePublic)
        
        .Verification = "Verification: CommComps.dat has been updated accordingly"
        .ResultExpected = CommonPublic.LastModAt("mBasic")
        .Result = CommonServiced.LastModAt("mBasic")
          
        '==========================================================================
        .TestId = "0100-3"
        .TestedComp = "clsServices"
        .TestedProc = "ExportChangedComponents"
        .TestedType = "Method"
        
        .Verification = "Precondition: ""mBasic.bas"" is the only export file in the export folder"
        .ResultExpected = True ' from the above update test
        .Result = FSo.GetFolder(mEnvironment.ExportServiceFolderPath).Files.Count = 1 And FSo.GetFolder(mEnvironment.ExportServiceFolderPath).Files("mBasic.bas").Name = "mBasic.bas"
        '-------------------------------
        Services.ExportChangedComponents
        '-------------------------------
        .Verification = "Verification: Exported number of components corresponds with the VBProject components"
        .ResultExpected = VBProjectExportFiles
        .Result = FSo.GetFolder(mEnvironment.ExportServiceFolderPath).Files.Count
        
        '==========================================================================
    End With
    Test_0100_2CleanUp
        
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_0100_2CleanUp()
    
    TestAid.CleanUp

End Sub

Private Sub Test_Timer()
    With New clsTestAid
        Debug.Print .TimerOverhead
    End With
End Sub

Private Sub Test_0010_Message()
    Dim Msg As mMsg.udtMsg
    
    Msg.Section(1).Text.Text = "Message test"
    mMsg.Dsply "Message test", Msg, , mMsg.Buttons("Ok", "Not Ok")
    
End Sub

Private Sub Test_0100_1SetUp()
    
    Dim sTestFolder As String
    
    With TestAid
'        .FolderZip "0100" ' !!! only when preparation had been changed !!!
        .FolderUnZip "Test_0100.zip", sTestFolder:                      .TempTestItem = sTestFolder ' indicate the result folder as a temporary test resuurce
        Set wbkServiced = Workbooks.Open(TestWorkbook(sTestFolder)):    .TempTestItem = wbkServiced ' indicate the Workbook as a temporary test resuource closed with CleanUp
    End With

End Sub

Private Sub Test_0200_2CleanUp()
    
    Dim v As Variant
    
    '~~ Remove any test component pending or public
    If mBasic.ArryIsAllocated(aTestComps) Then
        For Each v In aTestComps
            With CommonPending
                If .Exists(v) Then .Remove v
            End With
            With CommonPublic
                If .Exists(v) Then .Remove v
            End With
        Next v
    End If
    TestAid.CleanUp
    
End Sub

Private Sub Test_0200_1SetUp()

    Const PROC = "Test_0200_1SetUp"
            
    With TestAid
        .TestHeadLine = "Conflicts detected and handled by the Export service"
'        .FolderZip "0200" ' !!! only when preparation had been changed !!!
        .FolderUnZip "Test_0200.zip", sTestFolder:                      .TempTestItem = sTestFolder ' indicate the result folder as a temporary test resuurce
        sTestFolder = sTestFolder & "\Test_0200a"
        Set wbkServiced = Workbooks.Open(TestWorkbook(sTestFolder)):    .TempTestItem = wbkServiced ' indicate the Workbook as a temporary test resuource closed with CleanUp

        mCompMan.ServiceInitiate s_serviced_wbk:=wbkServiced _
                               , s_service:=TestAid.TestHeadLine _
                               , s_hosted:=sTestComp
        mExport.ChangedComponents
        .TempTestItem = CommonPending.LastModExpFile(sTestComp)
    End With
        
End Sub

Private Sub Test_0300_1SetUp_1()
' ------------------------------------------------------------------------------
' Note: this test is performed with ThisWorkbook directly.
' ------------------------------------------------------------------------------
    Const PROC = "Test_0300_1SetUp_1"
    
    With TestAid
        mCompMan.ServiceInitiate s_serviced_wbk:=ThisWorkbook _
                               , s_service:=TestAid.TestHeadLine
        With New clsComp
            .CompName = sTestComp
            FSo.CopyFile .ExpFileFullName, mEnvironment.CommCompsPath & "\" & sTestComp & ".cls"
        End With
    End With
        
End Sub

Private Sub Test_0300_1SetUp_2()
' ------------------------------------------------------------------------------
' Note: this test is performed with ThisWorkbook directly.
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_0300_1SetUp"
    
    sTestComp = "clsCode"
    
    mCompMan.ServiceInitiate s_serviced_wbk:=ThisWorkbook _
                           , s_service:=TestAid.TestHeadLine
    FSo.DeleteFile mEnvironment.CommCompsPath & "\" & sTestComp & ".cls"
        
End Sub

Private Property Let TestComp(ByVal t_comp As String)
    If sTestComp <> t_comp Then
        mBasic.Arry(aTestComps) = t_comp
        sTestComp = t_comp
    End If
End Property

Public Sub Test_0200_ConflictingExport()
' -----------------------------------------------------------------
' Headline: Conflicts detected and handled by the Export service.
' Covers: - Export confliction due to a code modification on a
'           non up-to-date Common Component
'         - Export conflicting due to an already pending release
'           due to a code modification made in another Workbook.
' Setup:
' ------
' Workbook       Common       State
'                Component
' -------------- ------------ ------------------------------------
' Test0200a.xlsb mBasic       public, outdated, modified
'                m0200Pending status hosted, pending release,
'                             yet no public version
' Test0290b.xlsb mBasic       public, outdated, modified
'                m0200Pending not up-to-date, modified, conflict
'                             with pending in Test0200a.xlsb
' -----------------------------------------------------------------
    Const PROC = "Test_0200_ConflictingExport"
    
    On Error GoTo eh
    Dim Comp As New clsComp
    Dim q    As clsQ
    
    On Error GoTo eh
    Prepare
    TestComp = "mTest0200Pending"
        
    With TestAid
        .TestId = "0200" ' basic test number for the environment setup
        .TestHeadLine = "Conflicts detected and handled by the Export service"
        Test_0200_1SetUp
        '~~ Assert precondition 1
        .Verification = "Precondition 1: The test component " & sTestComp & " is pending release"
        .ResultExpected = True
        .Result = CommonPending.Exists(sTestComp)
        
        '==========================================================================
        .TestId = "0200-1"
        sTestFolder = Replace(sTestFolder, "\Test_0200a", "\Test_0200b")
        ServicedWrkbk = Workbooks.Open(TestWorkbook(sTestFolder)):    .TempTestItem = ServicedWrkbk ' indicate the Workbook as a temporary test resuource closed with CleanUp

        mCompMan.ServiceInitiate s_serviced_wbk:=ServicedWrkbk _
                               , s_service:=TestAid.TestHeadLine _
                               , s_hosted:=sTestComp
        .RequiredInteraction = "Reply with ...."
        mExport.ChangedComponents sTestComp
        .ResultExpected = True
        .Result = True
             
    End With

xt: Test_0200_2CleanUp
    Exit Sub
    
eh: Test_0200_2CleanUp
    Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0300_CommCompManuallyCopiedRemoved()
' -----------------------------------------------------------------
' Headline: Common Component manually copied/removed in/from Common-Components folder.
' Covers: - Export confliction due to a code modification on a
'           non up-to-date Common Component
'         - Export conflicting due to an already pending release
'           due to a code modification made in another Workbook.
' Setup:
' ------
' Workbook       Common       State
'                Component
' -------------- ------------ ------------------------------------
' Test0300a.xlsb mBasic       public, outdated, modified
'                m0300Pending status hosted, pending release,
'                             yet no public version
' Test0290b.xlsb mBasic       public, outdated, modified
'                m0300Pending not up-to-date, modified, conflict
'                             with pending in Test0300a.xlsb
' -----------------------------------------------------------------
    Const PROC = "Test_0300_CommCompManuallyCopiedRemoved"
    
    On Error GoTo eh
    Dim Comp As New clsComp
    Dim q    As clsQ
    
    On Error GoTo eh
    Prepare
    sTestComp = "clsCode"
    
    With TestAid
        .TestId = "0300-1" ' basic test number for the environment setup
        .TestHeadLine = "Common Component manually copied/removed in/from Common-Components folder"
        .Verification = "Precondition 1: The test component " & sTestComp & " does not exist in the Common-Components folder"
        .ResultExpected = True
        With New clsCommonPublic
            TestAid.Result = Not FSo.FileExists(mEnvironment.CommCompsPath & "\" & sTestComp & ".cls")
        End With
        .Verification = "Precondition 2: The test component " & sTestComp & " does not exist in the serviced Workbooks CommComps.dat file"
        .ResultExpected = True
        With New clsCommonServiced
            TestAid.Result = Not .Components.Exists(sTestComp)
        End With
        Test_0300_1SetUp_1  ' copy Export-File into Common-Components folder and run housekeeping
        .Verification = "Precondition: The test component " & sTestComp & " exist in the Common-Components folder"
        .ResultExpected = True
        With New clsCommonPublic
            TestAid.Result = FSo.FileExists(mEnvironment.CommCompsPath & "\" & sTestComp & ".cls")
        End With
        mHskpng.CommComps
        .Verification = "Test result 1: Test component " & sTestComp & " registered as new Common Component"
        .ResultExpected = True
        .Result = CommonPublic.Exists(sTestComp)
        .Verification = "Test result 2: Test component " & sTestComp & " registered as ""used"" in CommComps.dat file"
        .ResultExpected = True
        .Result = CommonServiced.KindOfComponent(sTestComp) = enCompCommonUsed
        .Verification = "Test result 3: Test component " & sTestComp & " properties serviced equal public"
        .ResultExpected = True
        With New clsComp
            .CompName = sTestComp
            TestAid.Result = .ServicedMeetPublicProperties
        End With
        
        '==========================================================================
        .TestId = "0300-2"
        Test_0300_1SetUp_1  ' remove the test component from the Common-Components folder
        mHskpng.CommComps
        .Verification = "Test result 1: Test component " & sTestComp & " removed from CommComps.dat file"
        .ResultExpected = True
        .Result = Not CommonPublic.Components.Exists(sTestComp)
             
    End With

xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_0500_UpdateOutdatedCommonComponents()
    Const PROC  As String = "Test_0500_UpdateOutdatedCommonComponents"
    
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

Private Function VBProjectExportFiles() As Long
    Dim l   As Long
    Dim vbc As VBComponent
    
    For Each vbc In Serviced.Wrkbk.VBProject.VBComponents
        If vbc.Type = vbext_ct_MSForm Then l = l + 2 Else l = l + 1
    Next vbc
    VBProjectExportFiles = l
    
End Function

Sub AllEnvironVariables()
    Dim i           As Long
    
    For i = 1 To 255
        If Len(Environ$(i)) >= 3 Then
            Debug.Print "Environ(""" & Split(Environ$(i), "=")(0) & """) = " & Split(Environ$(i), "=")(1)
        End If
    Next i
    
End Sub
