Attribute VB_Name = "mCompManTest"
Option Explicit
Option Private Module
' -------------------------------------------------------
' Standard module mTest
'                 Provides the means to test the services
'                 of the CompMan AddIn Workbook.
'
' W. Rauschenberger Berlin, Jan 2021
' -------------------------------------------------------
Private Const TEST_CHANGE = "' Test code change"

Private cTest   As clsTestService
Private wbTest  As Workbook
Private wbSrc   As Workbook
Private wbTrgt  As Workbook
Private vbc     As VBComponent
Private vbcm    As CodeModule

Private Property Get mRenew_ByImport() As String
    mRenew_ByImport = CompManAddinName & "!mRenew.ByImport"
End Property

Public Sub ClearIW()
    Application.VBE.Windows("Direktbereich").SetFocus
        If Application.VBE.ActiveWindow.Caption = "Direktbereich" And Application.VBE.ActiveWindow.Visible Then
        Application.SendKeys "^a {DEL} {HOME}"
    End If
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest" & "." & sProc
End Function

Public Sub Regression()
' -----------------------------------------------------------------
'
' -----------------------------------------------------------------
    Const PROC = "Regression"
    
    Set cTest = New clsTestService
    cTest.Regression = True
    
    mBasic.BoP ErrSrc(PROC)
    Test_01_KindOfComp
    mErH.EoP ErrSrc(PROC)
    
End Sub

Public Sub RemoveTestCodeChange( _
                 Optional ByVal exp_file As String = vbNullString, _
                 Optional ByRef vbc As VBComponent = Nothing)
' ------------------------------------------------------------------
' Removes a code line from the provided VBComponent (vbc) which has
' been added for test purpose.
' Used to reset the test environment to its initial state
' ------------------------------------------------------------------
    Const PROC = "RemoveTestCodeChange"
    
    On Error GoTo eh
    Dim Comp   As clsComp
    
    If exp_file <> vbNullString Then
        With New FileSystemObject
            If .FileExists(exp_file) Then .DeleteFile exp_file
        End With
    End If
    
    If Not vbc Is Nothing Then
        With vbc.CodeModule
            If .Lines(1, 1) = TEST_CHANGE Then .DeleteLines 1, 1
            While Len(.Lines(1, 1)) = 0
                .DeleteLines 1.1
            Wend
        End With
    End If
    
    If Not Comp Is Nothing Then Set Comp = Nothing
    If Not vbcm Is Nothing Then Set vbcm = Nothing
    If Not vbc Is Nothing Then Set vbc = Nothing
    On Error Resume Next: wbTest.Close SaveChanges:=False
    On Error Resume Next: wbSrc.Close SaveChanges:=False
    On Error Resume Next: wbTrgt.Close SaveChanges:=False

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_01_KindOfComp()
    Const PROC = "Test_01_KindOfComp"

    Dim wb      As Workbook
    Dim fso     As New FileSystemObject
    Dim Comp   As clsComp
    Dim sComp   As String
    
    Set wb = mCompMan.WbkGetOpen(fso.GetParentFolderName(ThisWorkbook.Path) & "\File\File.xlsm")

    sComp = "mFile"
    Set Comp = Nothing
    Set Comp = New clsComp
    With Comp
        Set .Wrkbk = wb
        Set .VBComp = wb.VBProject.VBComponents(sComp)
        Debug.Assert .KindOfComp() = enKindOfComp.enCommCompUsed
    End With

    sComp = "fMsg"
    Set Comp = Nothing
    Set Comp = New clsComp
    With Comp
        Set .Wrkbk = wb
        Set .VBComp = wb.VBProject.VBComponents(sComp)
        Debug.Assert .KindOfComp() = enCommCompUsed
    End With
    
    sComp = "mTest"
    Set Comp = Nothing
    Set Comp = New clsComp
    With Comp
        Set .Wrkbk = wb
        Set .VBComp = wb.VBProject.VBComponents(sComp)
        Debug.Assert .KindOfComp() = enInternal
    End With
    
xt: wb.Close SaveChanges:=False
    Set Comp = Nothing
    Set fso = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_10_ExportChangedComponents()
    Const PROC = "Test_ExportChangedComponents"
    
    mBasic.BoP ErrSrc(PROC)
    mCompMan.ExportChangedComponents ThisWorkbook
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_BasicConfig()

    If Not mMe.BasicConfig Then
        Debug.Print "Basic configuration invalid!"
    End If
    
End Sub

Public Sub Test_Comps_Changed()
    
    Dim Comps   As New clsComps
    Dim Comp    As New clsComp
    Dim v       As Variant
    Dim dct     As Dictionary
    Dim Log     As New clsLog
    Dim s       As String
    
    Set mService.Serviced = ThisWorkbook
    Set Stats = New clsStats
    
    Set dct = Comps.Changed(Stats)
    For Each v In dct
        Stats.Count sic_used_comm_comps
        Set Comp = dct(v)
        s = s & Comp.CompName & ", "
        Set Comp = Nothing
    Next v
    
    With Stats
        Debug.Print .Total(sic_comps_total) & " Components"
        Debug.Print .Total(sic_comps_changed) & " Changed (" & Left(s, Len(s) - 2) & ")"
    End With
    Set Comps = Nothing
    Set Stats = Nothing
    Set Log = Nothing

End Sub

Public Sub Test_Comps_Outdated()
    
    Dim Comps   As New clsComps
    Dim Comp    As New clsComp
    Dim v       As Variant
    Dim dct     As Dictionary
    Dim Log     As New clsLog
    
    Set mService.Serviced = ThisWorkbook
    Set Stats = New clsStats
    Log.File = mFile.Temp(, ".log")
    
    Set dct = Comps.Outdated
    For Each v In dct
        Stats.Count sic_used_comm_comps
        Set Comp = dct(v)
        Log.ServicedItem = Comp.VBComp
        If Comp.RawChanged Then
            Stats.Count sic_comps_changed
            Debug.Print Comp.CompName & " is regarded outdated"
        End If
        Set Comp = Nothing
    Next v
    
    Debug.Print mFile.Txt(Log.LogFile)
    Debug.Print Stats.Total(sic_comps_total) & " Components"
    Debug.Print Stats.Total(sic_comps_changed) & " Outdated"
        
    Set Comps = Nothing
    Set Stats = Nothing
    With New FileSystemObject
        If .FileExists(Log.File) Then .DeleteFile (Log.File)
    End With
    Set Log = Nothing

End Sub

Public Sub Test_Log()
    Const PROC = "Test_Log"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    
    Set Log = New clsLog
    Set mService.Serviced = ThisWorkbook
    With Log
        .Service = ErrSrc(PROC)
        .ServicedItem = " <component-name> "
        .Entry = "Tested"
        mMsg.Box Title:="Test-Log:" _
               , Prompt:=mFile.Txt(ft_file:=.LogFile.Path) _
               , box_monospaced:=True
        If fso.FileExists(.LogFile.Path) Then fso.DeleteFile .LogFile.Path
    End With
    
xt: Set Log = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_Path()

    With New FileSystemObject
        Debug.Print .GetFile(mFile.Temp()).Path
    End With
End Sub

Public Sub Test_Refs()
    
    Dim ref As Reference

    For Each ref In ThisWorkbook.VBProject.References
        With ref
            If InStr(.Description, "Applications Extensibility") <> 0 Then
                Debug.Print .Description
                Debug.Print .GUID
                Debug.Print "Major=" & .Major
                Debug.Print "Minor=" & .Minor
                Debug.Print TypeName(.GUID)
            End If
            If InStr(.Description, "Scripting Runtime") <> 0 Then
                Debug.Print .Description
                Debug.Print .GUID
                Debug.Print "Major=" & .Major
                Debug.Print "Minor=" & .Minor
                Debug.Print TypeName(.GUID)
            End If
        End With
    Next ref
    
End Sub

Public Sub Test_RenewComp(ByVal rnc_exp_file_full_name, _
                          ByVal rnc_comp_name As String)
' --------------------------------------------------------
' This test procedure is exclusively initiated within the
' 'CompMan-Development-Instance-Workbook and executed by
' the 'CompMan-Addin' which needs to be open.
' When these conditions are not met a message is displayed
' --------------------------------------------------------
    Const PROC = "Test_RenewComp"
    
    Dim Comp        As New clsComp
    Dim wbActive    As Workbook
    Dim wbTemp      As Workbook
    
    If mMe.IsDevInstnc Then GoTo xt
    
    Set mService.Serviced = ThisWorkbook
    Log.File = mFile.Temp(, ".log")
    Log.Service = "Renew Component Test"
    
    With Comp
        .CompName = rnc_comp_name
        Log.ServicedItem = .VBComp
        
        If .Wrkbk Is ActiveWorkbook Then
            Set wbActive = ActiveWorkbook
            Set wbTemp = Workbooks.Add ' Activates a temporary Workbook
            Log.Entry = "Active Workbook de-activated by creating a temporary Workbook"
        End If
            
        mRenew.ByImport rn_wb:=.Wrkbk _
                      , rn_comp_name:=.CompName _
                      , rn_raw_exp_file_full_name:=rnc_exp_file_full_name
    End With
    
xt: If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
        Log.Entry = "Temporary created Workbook closed without save"
        Set wbTemp = Nothing
        If Not ActiveWorkbook Is wbActive Then
            wbActive.Activate
            Log.Entry = "De-activated Workbook '" & wbActive.name & "' re-activated"
            Set wbActive = Nothing
        Else
            Log.Entry = "Workbook '" & wbActive.name & "' re-activated by closing the temporary created Workbook"
        End If
    End If
    Set Comp = Nothing
    Set Log = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_RenewComp_0_Regression()
    Const PROC = ""
    
    On Error GoTo eh
    If mMe.IsAddinInstnc Then Exit Sub
    
    mBasic.EoP ErrSrc(PROC)
'    Test_RenewComp_1a_Standard_Module_ExpFile_Remote "mFile", repeat:=1
'    Test_RenewComp_1b_Standard_Module_ExpFile_Local "mFile", repeat:=1
'    Test_RenewComp_2_Class_Module_ExpFile_Local "clsLog", repeat:=2
'    Test_RenewComp_3a_UserForm_ExpFile_Local "fMsg", repeat:=1
    Test_RenewComp_3b_UserForm_ExpFile_Remote "fMsg", repeat:=1

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_RenewComp_1a_Standard_Module_ExpFile_Remote( _
            ByVal test_comp_name As String, _
   Optional ByVal repeat As Long = 1)
' ----------------------------------------------------------------------------
' This is a kind of "burn-in" test in order to prove that a Standard Module
' can be renewed by the re-import of an Export-File.
' The test asserts that a Workbook is able to renew its own VBA code provided
' it is not active when it is done.
' ----------------------------------------------------------------------------
    Const PROC          As String = "Test_RenewComp_1a_UserForm_ExpFile_Remote"
    
    On Error GoTo eh
    Dim Comp            As New clsComp
    Dim i               As Long
    Dim sExpFile        As String
    Dim flExport        As File
    
    If mMe.IsAddinInstnc Then Exit Sub
    If mMe.IsDevInstnc Then
        If mMe.CompManAddinIsOpen Then
            ' ---------------------------------
            '~~ Arguments for the Run:
            '~~ rc_exp_file_full_name As String
            '~~ rc_comp_name As String
            '~~ rc_wb As Workbook
            '~~ -------------------------------
            mBasic.BoP ErrSrc(PROC)
            With Comp
                Set .Wrkbk = ThisWorkbook
                .CompName = test_comp_name
                            
                '~~ ------------------------------------------------------
                '~~ Second test with the selection of a remote Export-File
                '~~ ------------------------------------------------------
                If mFile.Picked(p_title:="Select an Export-File for the renewal of the component '" & .CompName & "'!" _
                              , p_init_path:=mExport.ExpFileFolderPath(.Wrkbk) _
                              , p_filters:="Export File,*." & Comp.ExpFileExt _
                              , p_file:=flExport) _
                Then sExpFile = flExport.Path
                For i = 1 To repeat
                    Application.Run mRenew_ByImport _
                                  , .Wrkbk _
                                  , .CompName _
                                  , sExpFile
                Next i
            End With
            mBasic.EoP ErrSrc(PROC)
        End If
    End If
    
xt: Set Comp = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_RenewComp_1b_Standard_Module_ExpFile_Local( _
            ByVal test_comp_name As String, _
   Optional ByVal repeat As Long = 1)
' ----------------------------------------------------------------------------
' This is a kind of "burn-in" test in order to prove that a Standard Module
' can be renewed by the re-import of an Export-File.
' The test asserts that a Workbook is able to renew its own VBA code provided
' it is not active when it is done.
' ----------------------------------------------------------------------------
    Const PROC = "Test_RenewComp_1b_Standard_Module_ExpFile_Local"
    
    On Error GoTo eh
    Dim Comp    As New clsComp
    Dim i       As Long
    
    If mMe.IsAddinInstnc Then Exit Sub
    If mMe.IsDevInstnc Then
        If mMe.CompManAddinIsOpen Then
            ' ---------------------------------
            '~~ Arguments for the Run:
            '~~ rc_exp_file_full_name As String
            '~~ rc_comp_name As String
            '~~ rc_wb As Workbook
            '~~ -------------------------------
            mBasic.BoP ErrSrc(PROC)
            With Comp
                Set .Wrkbk = ThisWorkbook
                .CompName = test_comp_name
            
                For i = 1 To repeat
                    Application.Run mRenew_ByImport _
                                  , .Wrkbk _
                                  , .CompName _
                                  , .ExpFileFullName
                Next i
            End With
            mBasic.EoP ErrSrc(PROC)
        End If
    End If
    
xt: Set Comp = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_RenewComp_2_Class_Module_ExpFile_Local( _
            ByVal test_comp_name As String, _
   Optional ByVal repeat As Long = 1)
' ----------------------------------------------------------------------------
' This is a kind of "burn-in" test in order to prove that a Standard Module
' can be renewed by the re-import of an Export-File.
' The test asserts that a Workbook is able to renew its own VBA code provided
' it is not active when it is done.
' ----------------------------------------------------------------------------
    Const PROC = "Test_RenewComp_2_Class_Module_ExpFile_Local"
    
    On Error GoTo eh
    Dim Comp    As New clsComp
    Dim i       As Long
    
    If mMe.IsAddinInstnc Then Exit Sub
    If mMe.IsDevInstnc Then
        If mMe.CompManAddinIsOpen Then
            ' ---------------------------------
            '~~ Arguments for the Run:
            '~~ rc_exp_file_full_name As String
            '~~ rc_comp_name As String
            '~~ rc_wb As Workbook
            '~~ -------------------------------
            mBasic.BoP ErrSrc(PROC)
            With Comp
                Set .Wrkbk = ThisWorkbook
                .CompName = test_comp_name
            
                For i = 1 To repeat
                    Application.Run mRenew_ByImport _
                                  , .Wrkbk _
                                  , .CompName _
                                  , .ExpFileFullName
                Next i
            End With
            mBasic.EoP ErrSrc(PROC)
        End If
    End If
    
xt: Set Comp = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_RenewComp_3a_UserForm_ExpFile_Local( _
            ByVal test_comp_name As String, _
   Optional ByVal repeat As Long = 1)
' ----------------------------------------------------------------------------
' This is a kind of "burn-in" test in order to prove that a Standard Module
' can be renewed by the re-import of an Export-File.
' The test asserts that a Workbook is able to renew its own VBA code provided
' it is not active when it is done.
' ----------------------------------------------------------------------------
    Const PROC          As String = "Test_RenewComp_3a_UserForm_ExpFile_Local"
    
    On Error GoTo eh
    Dim Comp    As New clsComp
    Dim i       As Long
    
    If mMe.IsAddinInstnc Then Exit Sub
    If mMe.IsDevInstnc Then
        If mMe.CompManAddinIsOpen Then
            ' ---------------------------------
            '~~ Arguments for the Run:
            '~~ rc_exp_file_full_name As String
            '~~ rc_comp_name As String
            '~~ rc_wb As Workbook
            '~~ -------------------------------
            mBasic.BoP ErrSrc(PROC)
            With Comp
                Set .Wrkbk = ThisWorkbook
                .CompName = test_comp_name
            
                '~~ -------------------------------------------------
                '~~ First test with the components origin Export-File
                '~~ -------------------------------------------------
                For i = 1 To repeat
                    Application.Run mRenew_ByImport _
                                  , .Wrkbk _
                                  , .CompName _
                                  , .ExpFileFullName
                Next i
            End With
        End If
    End If
    
xt: Set Comp = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_RenewComp_3b_UserForm_ExpFile_Remote( _
            ByVal test_comp_name As String, _
   Optional ByVal repeat As Long = 1)
' ----------------------------------------------------------------------------
' This is a kind of "burn-in" test in order to prove that a Standard Module
' can be renewed by the re-import of an Export-File.
' The test asserts that a Workbook is able to renew its own VBA code provided
' it is not active when it is done.
' ----------------------------------------------------------------------------
    Const PROC          As String = "Test_RenewComp_3b_UserForm_ExpFile_Remote"
    
    On Error GoTo eh
    Dim Comp        As New clsComp
    Dim i           As Long
    Dim sExpFile    As String
    Dim flExport    As File
    
    If mMe.IsAddinInstnc Then Exit Sub
    If mMe.IsDevInstnc Then
        If mMe.CompManAddinIsOpen Then
            ' ---------------------------------
            '~~ Arguments for the Run:
            '~~ rc_exp_file_full_name As String
            '~~ rc_comp_name As String
            '~~ rc_wb As Workbook
            '~~ -------------------------------
            mBasic.BoP ErrSrc(PROC)
            With Comp
                Set .Wrkbk = ThisWorkbook
                .CompName = test_comp_name
                            
                '~~ ------------------------------------------------------
                '~~ Second test with the selection of a remote Export-File
                '~~ ------------------------------------------------------
                If mFile.Picked(p_title:="Select an Export-File for the renewal of the component '" & .CompName & "'!" _
                              , p_init_path:=mExport.ExpFileFolderPath(.Wrkbk) _
                              , p_filters:="UserForm, *." & Comp.ExpFileExt _
                              , p_file:=flExport) _
                Then sExpFile = flExport.Path
                For i = 1 To repeat
                    Application.Run mRenew_ByImport _
                                  , .Wrkbk _
                                  , .CompName _
                                  , sExpFile
                Next i
            End With
            mBasic.EoP ErrSrc(PROC)
        End If
    End If
    
xt: Set Comp = Nothing
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_SheetControls_Name_and_Type()
' --------------------------------------------------
' List all sheet controls ordered ascending by name.
' --------------------------------------------------
    Const PROC = "Test_SheetControls_Name_and_Type"
    
    On Error GoTo eh
    Dim ws  As Worksheet
    Dim shp As Shape
    Dim dct As New Dictionary
    Dim v   As Variant
    Dim sWb As String
    Dim sWs As String
    
    sWb = "E:\Ablage\Excel VBA\DevAndTest\Excel-VB-Project-Component-Management-Services\Test\SyncSource\SyncSource.xlsb"
    sWs = "Test_B1"
    
    Set ws = mCompMan.WbkGetOpen(sWb).Worksheets(sWs)
    For Each shp In ws.Shapes
        mDct.DctAdd dct, mSheetControls.CntrlName(shp), ws.name & "(" & mSheetControls.CntrlType(shp) & ")", order_bykey, seq_ascending, sense_caseignored, , True
    Next shp
    For Each v In dct
        Debug.Print dct(v), Tab(45), v
    Next v

xt: Set dct = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_UpdateOutdatedCommonComponents()
    Const PROC  As String = "Test_UpdateOutdatedCommonComponents"
    
    On Error GoTo eh
    Dim AddinService    As String
    Dim AddinStatus     As String
    
    If mService.Denied Then GoTo xt

    AddinService = CompManAddinName & "!mCompMan.UpdateOutdatedCommonComponents"
    If mMe.CompManAddinIsOpen Then
        AddinStatus = " (currently the case) "
    Else
        AddinStatus = " (currently  " & mBasic.Spaced("not") & "  the case) "
    End If
    
    If mMe.IsDevInstnc Then
        mBasic.BoP ErrSrc(PROC)
        
        On Error Resume Next
        Application.Run AddinService _
                      , ThisWorkbook
        
        If Err.Number = 1004 Then
            MsgBox Title:="CompMan-Addin not open (required for test: " & PROC & "!" _
                 , Prompt:="Application.Run " & vbLf & vbLf & AddinService & vbLf & vbLf & "failed because the 'CompMan-Addin' is not open!" _
                 , Buttons:=vbExclamation
        End If
        mBasic.EoP ErrSrc(PROC)
    Else
        MsgBox Title:="Test " & PROC & " not executed!" _
             , Prompt:="Executions of this test must not be performed 'within' the 'CompMan-Addin' Workbook." & vbLf & vbLf & _
                       "The test requires the 'CompMan-Addin' (" & mMe.CompManAddinName & ") is open " & AddinStatus & " but must be performed " & _
                       "from within the development instance (" & mMe.DevInstncFullName & ")." _
             , Buttons:=vbExclamation
    End If
    
xt: Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_WinMerge()

    Dim sF1 As String
    Dim sF2 As String
    Dim f1  As File
    Dim f2  As File
    
    mFile.Picked p_file:=f1
    mFile.Picked p_file:=f2
    sF1 = f1.Path
    sF2 = f2.Path
    
    mService.ExpFilesDiffDisplay fd_exp_file_left_full_name:=sF1 _
                                   , fd_exp_file_right_full_name:=sF2 _
                                   , fd_exp_file_left_title:=sF1 _
                                   , fd_exp_file_right_title:=sF2

End Sub

