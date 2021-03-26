Attribute VB_Name = "mTest"
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
Private cComp   As New clsComp
Private wbTest  As Workbook
Private wbSrc   As Workbook
Private wbTrgt  As Workbook
Private vbc     As VBComponent
Private vbcm    As CodeModule

Private Property Get mRenew_ByImport() As String:           mRenew_ByImport = CompManAddinName & "!mRenew.ByImport":                    End Property

Private Property Get mService_UpdateRawClones() As String:  mService_UpdateRawClones = CompManAddinName & "!mService.UpdateRawClones":  End Property

Public Sub Cleanup(Optional ByVal exp_file As String = vbNullString, _
                    Optional ByRef vbc As VBComponent = Nothing)
        
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
    
    If Not cComp Is Nothing Then Set cComp = Nothing
    If Not vbcm Is Nothing Then Set vbcm = Nothing
    If Not vbc Is Nothing Then Set vbc = Nothing
    On Error Resume Next: wbTest.Close SaveChanges:=False
    On Error Resume Next: wbSrc.Close SaveChanges:=False
    On Error Resume Next: wbTrgt.Close SaveChanges:=False

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest" & "." & sProc
End Function

Private Function MaxCompLength(ByRef wb As Workbook) As Long
    Dim vbc As VBComponent
    If lMaxCompLength = 0 Then
        For Each vbc In wb.VBProject.VBComponents
            MaxCompLength = mBasic.Max(MaxCompLength, Len(vbc.Name))
        Next vbc
    End If
End Function

Public Sub Regression()
' -----------------------------------------------------------------
'
' -----------------------------------------------------------------
    Const PROC = "Regression"
    
    Set cTest = New clsTestService
    cTest.Regression = True
    
    mErH.BoP ErrSrc(PROC)
    Test_01_KindOfComp
    mErH.EoP ErrSrc(PROC)
    
End Sub
  
Public Sub Test_SynchTargetWithSource()
' ---------------------------------------------------------------------
' Attention: This test preserves the target Workbook by a backup before
' and a restore after the synch test. The target Workbook thus will not
' show the synch result unless the terst procedire is stopped.
' ---------------------------------------------------------------------
    Const PROC = "Test_SynchTargetWithSource"
    
    On Error GoTo eh
    Dim sSource     As String
    Dim sTarget     As String
    Dim BkpFolder   As String
    
    sTarget = "E:\Ablage\Excel VBA\DevAndTest\Test-Sync-Target-Project\Test_Sync_Target.xlsb"
    sSource = "E:\Ablage\Excel VBA\DevAndTest\Test-Sync-Source-Project\Test_Sync_Source.xlsb"
       
    If mService.SyncTargetWithSource(wb_target:=mCompMan.WbkGetOpen(sTarget) _
                                   , wb_source_name:=sSource _
                                   , restricted_sheet_rename_asserted:=True _
                                   , bkp_folder:=BkpFolder _
                                    ) _
    Then
        ' -----------------------------------------------------------------
        Stop ' last chance to see the synch result in the target Workbook !
        '    ' traget Workbook will be restored as test Workbook
        ' -----------------------------------------------------------------
    End If
    '~~ The backup is no longer reqired
    
xt:
    If BkpFolder <> vbNullString Then
        ' A backup had been made by the service which is now restored and the target is re-opened
        mCompMan.WbkGetOpen(sTarget).Close SaveChanges:=False
        mSync.SyncRestore BkpFolder
        Application.EnableEvents = False ' The open service UpdateRawClones would start with a new log-file otherwise
        mCompMan.WbkGetOpen sTarget
        Application.EnableEvents = True
    End If
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub Test_01_KindOfComp()
    Const PROC = "Test_01_KindOfComp"

    Dim wb      As Workbook
    Dim fso     As New FileSystemObject
    Dim cComp   As clsComp
    Dim sComp   As String
    
    Set wb = mCompMan.WbkGetOpen(fso.GetParentFolderName(ThisWorkbook.Path) & "\File\File.xlsm")

    sComp = "mFile"
    Set cComp = Nothing
    Set cComp = New clsComp
    With cComp
        Set .Wrkbk = wb
        Set .VBComp = wb.VBProject.VBComponents(sComp)
        Debug.Assert .KindOfComp() = enKindOfComp.enRawClone
    End With

    sComp = "fMsg"
    Set cComp = Nothing
    Set cComp = New clsComp
    With cComp
        Set .Wrkbk = wb
        Set .VBComp = wb.VBProject.VBComponents(sComp)
        Debug.Assert .KindOfComp() = enRawClone
    End With
    
    sComp = "mTest"
    Set cComp = Nothing
    Set cComp = New clsComp
    With cComp
        Set .Wrkbk = wb
        Set .VBComp = wb.VBProject.VBComponents(sComp)
        Debug.Assert .KindOfComp() = enInternal
    End With
    
xt: wb.Close SaveChanges:=False
    Set cComp = Nothing
    Set fso = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_10_ExportChangedComponents()
    Const PROC = "Test_ExportChangedComponents"
    
    mErH.BoP ErrSrc(PROC)
    mCompMan.ExportChangedComponents ThisWorkbook
    mErH.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_Log()
    Const PROC = "Test_Log"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim cLog        As New clsLog
    
    With cLog
        Set .ServicedWrkbk(sw_new_log:=True) = ThisWorkbook
        .Service = ErrSrc(PROC)
        .ServicedItem = " <component-name> "
        .Entry = "Tested"
        mMsg.Box msg_title:="Test-Log:" _
               , msg:=mFile.Txt(ft_file:=.LogFile.Path) _
               , msg_monospaced:=True
        If fso.FileExists(.LogFile.Path) Then fso.DeleteFile .LogFile.Path
    End With
    
xt: Set Log = Nothing
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select

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
    
    Dim cLog        As New clsLog
    Dim cComp       As New clsComp
    Dim wbActive    As Workbook
    Dim wbTemp      As Workbook
    
    If mMe.IsDevInstnc Then GoTo xt
    
    Log.File = mFile.Temp(, ".log")
    Set Log.ServicedWrkbk(sw_new_log:=True) = ThisWorkbook
    Log.Service = PROC
    
    With cComp
        .CompName = rnc_comp_name
        Log.ServicedItem = .VBComp
        
        If .Wrkbk Is ActiveWorkbook Then
            Set wbActive = ActiveWorkbook
            Set wbTemp = Workbooks.Add ' Activates a temporary Workbook
            Log.Entry = "Active Workbook de-activated by creating a temporary Workbook"
        End If
            
        mRenew.ByImport rn_wb:=.Wrkbk _
                      , rn_comp_name:=.CompName _
                      , rn_exp_file_full_name:=rnc_exp_file_full_name
    End With
    
xt: If Not wbTemp Is Nothing Then
        wbTemp.Close SaveChanges:=False
        Log.Entry = "Temporary created Workbook closed without save"
        Set wbTemp = Nothing
        If Not ActiveWorkbook Is wbActive Then
            wbActive.Activate
            Log.Entry = "De-activated Workbook '" & wbActive.Name & "' re-activated"
            Set wbActive = Nothing
        Else
            Log.Entry = "Workbook '" & wbActive.Name & "' re-activated by closing the temporary created Workbook"
        End If
    End If
    Set cComp = Nothing
    Set Log = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: GoTo xt
    End Select
End Sub

Public Sub Test_RenewComp_0_Regression()
    Const PROC = ""
    
    On Error GoTo eh
    If mMe.IsAddinInstnc Then Exit Sub
    
    mErH.EoP ErrSrc(PROC)
'    Test_RenewComp_1a_Standard_Module_ExpFile_Remote "mFile", repeat:=1
'    Test_RenewComp_1b_Standard_Module_ExpFile_Local "mFile", repeat:=1
'    Test_RenewComp_2_Class_Module_ExpFile_Local "clsLog", repeat:=2
'    Test_RenewComp_3a_UserForm_ExpFile_Local "fMsg", repeat:=1
    Test_RenewComp_3b_UserForm_ExpFile_Remote "fMsg", repeat:=1

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
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
    Dim cComp           As New clsComp
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
            mErH.BoP ErrSrc(PROC)
            With cComp
                Set .Wrkbk = ThisWorkbook
                .CompName = test_comp_name
                            
                '~~ ------------------------------------------------------
                '~~ Second test with the selection of a remote Export-File
                '~~ ------------------------------------------------------
                If mFile.SelectFile(sel_init_path:=mCompMan.ExpFileFolderPath(.Wrkbk) _
                                  , sel_filters:="*" & cComp.ExpFileExt _
                                  , sel_filter_name:="bas-Export-Files" _
                                  , sel_title:="Select an Export-File for the renewal of the component '" & .CompName & "'!" _
                                  , sel_result:=flExport) _
                Then sExpFile = flExport.Path
                For i = 1 To repeat
                    Application.Run mRenew_ByImport _
                                  , .Wrkbk _
                                  , .CompName _
                                  , sExpFile
                Next i
            End With
            mErH.EoP ErrSrc(PROC)
        End If
    End If
    
xt: Set cComp = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
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
    Dim cComp   As New clsComp
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
            mErH.BoP ErrSrc(PROC)
            With cComp
                Set .Wrkbk = ThisWorkbook
                .CompName = test_comp_name
            
                For i = 1 To repeat
                    Application.Run mRenew_ByImport _
                                  , .Wrkbk _
                                  , .CompName _
                                  , .ExpFileFullName
                Next i
            End With
            mErH.EoP ErrSrc(PROC)
        End If
    End If
    
xt: Set cComp = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
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
    Dim cComp   As New clsComp
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
            mErH.BoP ErrSrc(PROC)
            With cComp
                Set .Wrkbk = ThisWorkbook
                .CompName = test_comp_name
            
                For i = 1 To repeat
                    Application.Run mRenew_ByImport _
                                  , .Wrkbk _
                                  , .CompName _
                                  , .ExpFileFullName
                Next i
            End With
            mErH.EoP ErrSrc(PROC)
        End If
    End If
    
xt: Set cComp = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
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
    Dim cComp           As New clsComp
    Dim i               As Long
    
    If mMe.IsAddinInstnc Then Exit Sub
    If mMe.IsDevInstnc Then
        If mMe.CompManAddinIsOpen Then
            ' ---------------------------------
            '~~ Arguments for the Run:
            '~~ rc_exp_file_full_name As String
            '~~ rc_comp_name As String
            '~~ rc_wb As Workbook
            '~~ -------------------------------
            mErH.BoP ErrSrc(PROC)
            With cComp
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
    
xt: Set cComp = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
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
    Dim cComp           As New clsComp
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
            mErH.BoP ErrSrc(PROC)
            With cComp
                Set .Wrkbk = ThisWorkbook
                .CompName = test_comp_name
                            
                '~~ ------------------------------------------------------
                '~~ Second test with the selection of a remote Export-File
                '~~ ------------------------------------------------------
                If mFile.SelectFile(sel_init_path:=mCompMan.ExpFileFolderPath(.Wrkbk) _
                                  , sel_filters:="*" & cComp.ExpFileExt _
                                  , sel_filter_name:="UserForm" _
                                  , sel_title:="Select an Export-File for the renewal of the component '" & .CompName & "'!" _
                                  , sel_result:=flExport) _
                Then sExpFile = flExport.Path
                For i = 1 To repeat
                    Application.Run mRenew_ByImport _
                                  , .Wrkbk _
                                  , .CompName _
                                  , sExpFile
                Next i
            End With
            mErH.EoP ErrSrc(PROC)
        End If
    End If
    
xt: Set cComp = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_UpdateRawClones()
    Const PROC  As String = "Test_UpdateRawClones"
    
    On Error GoTo eh
    Dim AddinService    As String
    Dim AddinStatus     As String
    
    If mService.Denied(den_serviced_wb:=ThisWorkbook, den_service:=PROC) Then GoTo xt

    AddinService = CompManAddinName & "!mService.UpdateRawClones"
    If mMe.CompManAddinIsOpen Then
        AddinStatus = " (currently the case) "
    Else
        AddinStatus = " (currently  " & mBasic.Spaced("not") & "  the case) "
    End If
    
    If mMe.IsDevInstnc Then
        mErH.BoP ErrSrc(PROC)
        
        On Error Resume Next
        Application.Run AddinService _
                      , ThisWorkbook
        
        If Err.Number = 1004 Then
            MsgBox Title:="CompMan-Addin not open (required for test: " & PROC & "!" _
                 , Prompt:="Application.Run " & vbLf & vbLf & AddinService & vbLf & vbLf & "failed because the 'CompMan-Addin' is not open!" _
                 , Buttons:=vbExclamation
        End If
        mErH.EoP ErrSrc(PROC)
    Else
        MsgBox Title:="Test " & PROC & " not executed!" _
             , Prompt:="Executions of this test must not be performed 'within' the 'CompMan-Addin' Workbook." & vbLf & vbLf & _
                       "The test requires the 'CompMan-Addin' (" & mMe.CompManAddinName & ") is open " & AddinStatus & " but must be performed " & _
                       "from within the development instance (" & mMe.DevInstncFullName & ")." _
             , Buttons:=vbExclamation
    End If
    
xt: Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub

Public Sub Test_BasicConfig()

    If Not mMe.BasicConfig Then
        Debug.Print "Basic configuration invalid!"
    End If
    
End Sub

Public Sub Test_Changed_Clones()
    
    Dim Clones      As New clsClones
    
    Set Stats = New clsStats
    Set Clones.Serviced = ThisWorkbook
    Set Log.ServicedWrkbk(True) = ThisWorkbook
    Log.File = mFile.Temp(, ".log")
    Clones.CollectAllChanged
    
    Debug.Print Stats.Total(sic_comps_total) & " Components"
    Debug.Print Stats.Total(sic_clone_comps) & " Clones"
    Debug.Print Stats.Total(sic_clone_changed) & " Changed"
    Debug.Print mFile.Txt(Log.LogFile)
    
    Set Clones = Nothing
    Set Stats = Nothing
    With New FileSystemObject
        If .FileExists(Log.File) Then .DeleteFile (Log.File)
    End With
    Set Log = Nothing
End Sub

Public Sub Test_Changed_Comps()
    
    Dim Comps   As New clsComps
    Dim Comp    As New clsComp
    Dim v       As Variant
    
    Set Stats = New clsStats
    Set Comps.Serviced = ThisWorkbook
    Set Log.ServicedWrkbk(True) = ThisWorkbook
    Log.File = mFile.Temp(, ".log")
    Comps.CollectAllChanged
    
    For Each v In Comps.Changed
        Set Comp = Comps.Changed(v)
        Log.ServicedItem = Comp.VBComp
        Log.Entry = "Code changed (export due)"
        Set Comp = Nothing
    Next v
    
    Debug.Print mFile.Txt(Log.LogFile)
    Debug.Print Stats.Total(sic_comps_total) & " Components"
    Debug.Print Stats.Total(sic_comps_changed) & " Changed"
        
    Set Comps = Nothing
    Set Stats = Nothing
    With New FileSystemObject
        If .FileExists(Log.File) Then .DeleteFile (Log.File)
    End With
    Set Log = Nothing

End Sub

Public Sub Test_SynchFormating()
    Const PROC = "Test_SynchFormating"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim BkpFolder   As String
    Dim sTarget     As String
    Dim sSource     As String
    Dim BttnDelete  As String
    Dim BttnKeep    As String
    Dim flLog       As File
    Dim LastModif   As Date
    
    Set Log = New clsLog
    Log.File = mFile.Temp(ThisWorkbook.Path, ".log")
    
    Set Stats = New clsStats
    Set Sync = New clsSync
    
    sTarget = "E:\Ablage\Excel VBA\DevAndTest\Test-Sync-Target-Project\Test_Sync_Target.xlsb"
    sSource = "E:\Ablage\Excel VBA\DevAndTest\Test-Sync-Source-Project\Test_Sync_Source.xlsb"
    
    Set Sync.Target = mCompMan.WbkGetOpen(sTarget)
    Set Sync.Source = mCompMan.WbkGetOpen(sSource)
    Set Log.ServicedWrkbk = Sync.Target
    
    Sync.CollectAllSyncItems
    
    LastModif = fso.GetFile(sTarget).DateLastModified
    
    mSync.SyncBackup BkpFolder, sTarget
    mSyncRanges.SyncFormating
    
    If BkpFolder <> vbNullString Then
        ' A backup had been made by the service which is now restored and the target is re-opened
        
        mCompMan.WbkGetOpen(sTarget).Close SaveChanges:=False
        mSync.SyncRestore BkpFolder
        Application.EnableEvents = False ' The open service UpdateRawClones would start with a new log-file otherwise
        mCompMan.WbkGetOpen sTarget
        Application.EnableEvents = True
    End If

xt: With New FileSystemObject
        If .FileExists(Log.File) Then
            Set flLog = .GetFile(Log.File)
            BttnDelete = "Delete Log-File" & vbLf & .GetFileName(Log.File)
            BttnKeep = "Keep Log-File" & vbLf & .GetFileName(Log.File)
            If mMsg.Box(msg_title:=PROC & " Log-File" _
                      , msg:=mFile.Txt(.GetFile(Log.File)) _
                      , msg_monospaced:=True _
                      , msg_buttons:=mMsg.Buttons(BttnDelete, BttnKeep) = BttnDelete) _
            Then .DeleteFile flLog
        End If
    End With
    Set Log = Nothing

    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOptResumeErrorLine: Stop: Resume
        Case mErH.DebugOptResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub