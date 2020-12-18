Attribute VB_Name = "mCfg"
Option Explicit
Option Private Module

Private Const SECTION_BASE_CONFIG   As String = "BaseConfiguration"
Private Const COMMON_BASE_PATH      As String = "CommCompsBasePath"
Private Const COMPMAN_ADDIN_PATH    As String = "AddInPath"
Private Const HOSTED_FILE_NAME      As String = "HostedFileName"

Private Property Get CFG_CHANGE_ADDIN_PATH() As String
    CFG_CHANGE_ADDIN_PATH = "Change AddIn Path:" & vbLf & vbLf & "Select another folder"
End Property

Private Property Get CFG_CHANGE_COMM_COMPS_PATH() As String
    CFG_CHANGE_COMM_COMPS_PATH = "Change Common Component Workbooks Path:" & vbLf & vbLf & "Select another root folder"
End Property

Private Property Get CFG_CHANGE_HOSTED_FILENAME() As String
     CFG_CHANGE_HOSTED_FILENAME = "Change the name of the ""hosted"" files:" & vbLf & vbLf & "Select another ""hosted"" file example"
End Property

Private Property Get CFG_FILE_NAME() As String: CFG_FILE_NAME = ThisWorkbook.PATH & "\CompMan.cfg": End Property

Public Property Get CompManAddinPath() As String
    CompManAddinPath = mFl.Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=COMPMAN_ADDIN_PATH, vl_file:=CFG_FILE_NAME)
End Property

Public Property Let CompManAddinPath(ByVal s As String)
    mFl.Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=COMPMAN_ADDIN_PATH, vl_file:=CFG_FILE_NAME) = s
End Property

Public Property Get HostedCommCompsFileName() As String
    HostedCommCompsFileName = mFl.Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=HOSTED_FILE_NAME, vl_file:=CFG_FILE_NAME)
End Property

Public Property Let HostedCommCompsFileName(ByVal s As String)
    mFl.Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=HOSTED_FILE_NAME, vl_file:=CFG_FILE_NAME) = s
End Property

Public Property Get CommonComponentsBasePath() As String
    CommonComponentsBasePath = mFl.Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=COMMON_BASE_PATH, vl_file:=CFG_FILE_NAME)
End Property

Public Property Let CommonComponentsBasePath(ByVal s As String)
    mFl.Value(vl_section:=SECTION_BASE_CONFIG, vl_value_name:=COMMON_BASE_PATH, vl_file:=CFG_FILE_NAME) = s
End Property

Public Function Asserted() As Boolean
' ----------------------------------------------------
' Assert that an existing Common folder is configured
' and that it contains a subfolder "CommComponents".
' Attention! This function must not run in the AddIn
' instance of this Workbook!
' ----------------------------------------------------
    Const PROC = "Assert"
    
    On Error GoTo eh
    Dim sPathCommon     As String
    Dim sPathCompMan    As String
    Dim sHostedName     As String
    Dim v               As Variant
    Dim fl              As FILE
    
    With New FileSystemObject
        If .FolderExists(mCfg.CompManAddinPath) _
        And .FolderExists(mCfg.CommonComponentsBasePath) _
        And mFl.Exists(xst_file:=mCfg.CommonComponentsBasePath & "\" & mCfg.HostedCommCompsFileName & "*") _
        Then
            Asserted = True
            .CopyFile Source:=CFG_FILE_NAME, Destination:=mCfg.CompManAddinPath & "\CompMan.cfg", OverWriteFiles:=True
            GoTo xt
        End If
                
        '~~ Assert the folder for the CompMan AddIn
        sPathCompMan = mCfg.CompManAddinPath
        If sPathCompMan = vbNullString Then
            sPathCompMan = mBasic.SelectFolder( _
                           sTitle:="Select the folder for the AddIn instance of the CompManDev Workbook (escape to use the Application.UserLibraryPath)")
            If sPathCompMan = vbNullString Then
                sPathCompMan = Application.UserLibraryPath ' Default because user escaped the selection
            Else
                '~~ Assure trust in this location and save it to the CompMan.cfg file
                mDat.TrustThisFolder FolderPath:=sPathCompMan, TrustSubfolders:=False
                mCfg.CompManAddinPath = sPathCompMan
            End If
        Else
            While Not .FolderExists(sPathCompMan)
                '~~ Configured but folder not or no longer exists
                sPathCompMan = mBasic.SelectFolder( _
                               sTitle:="The configured CompMan AddIn folder does not exist. Select another one or escape for the default '" & Application.UserLibraryPath & "' path")
                If sPathCompMan = vbNullString Then
                    sPathCompMan = Application.UserLibraryPath
                Else
                    '~~ Assure trust in this location and save it to the CompMan.cfg file
                    mDat.TrustThisFolder FolderPath:=sPathCompMan, TrustSubfolders:=False
                    mCfg.CompManAddinPath = sPathCompMan
                End If
            Wend
        End If
                   
        '~~ Assert the Common Workbooks path
        sPathCommon = mCfg.CommonComponentsBasePath
        If sPathCommon = vbNullString Then
            sPathCommon = mBasic.SelectFolder("Select the root folder for the ""Common Component Workbooks""")
            mCfg.CommonComponentsBasePath = sPathCommon
        Else
            While Not .FolderExists(sPathCommon)
                sPathCommon = mBasic.SelectFolder( _
                sTitle:="Then current configured Common Component Workbook path does not exist. Select another one.")
            Wend
            If mCfg.CommonComponentsBasePath <> sPathCommon Then mCfg.CommonComponentsBasePath = sPathCommon
        End If
        
        '~~ Assert the name for the Hosted Common Components file name
        sHostedName = mCfg.HostedCommCompsFileName
        If sHostedName = vbNullString Then
            If mFl.SelectFile( _
                                sel_init_path:=mCfg.CommonComponentsBasePath _
                              , sel_title:="Select a file which - for example - is one indicating a Common Component(s) hosted in the coresponding Workbook." _
                              , sel_result:=fl _
                              ) _
            Then sHostedName = fl.ShortName
            mCfg.HostedCommCompsFileName = sHostedName
        End If
    
    End With

xt: Exit Function

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Function

Public Sub Confirm()
    Const PROC = "Confirm"
    Const CFG_CONFIRMED = "Confirmed"
    
    On Error GoTo eh
    Dim sMsg            As tMsg
    Dim sReply          As String
    Dim sPathCompMan    As String
    Dim sPathCommon     As String
    Dim fl              As FILE
    
    With sMsg
        .section(1).sLabel = "Location (path) for the AddIn instance of the CompManDev Workbook:"
        .section(1).sText = mCfg.CompManAddinPath & vbLf & _
                            "(when the selection is returned with no folder selected the path will default to """ & Application.UserLibraryPath & """)"
        .section(1).bMonspaced = True
        .section(2).sLabel = "Root folder for Common Component Workbooks:"
        .section(2).sText = mCfg.CommonComponentsBasePath
        .section(2).bMonspaced = True
        .section(3).sLabel = "Name of the ""hosted"" files (those which contains a Common Component Workbook's ""hosted"" Components):"
        .section(3).sText = mCfg.HostedCommCompsFileName
        .section(3).bMonspaced = True
    End With
    
    While sReply <> CFG_CONFIRMED
        sReply = mMsg.Dsply(msg_title:="Confirm or change the current Component Management's basic configuration" _
                          , msg_sections:=sMsg _
                          , msg_buttons:=mMsg.Buttons(CFG_CONFIRMED, vbLf, CFG_CHANGE_ADDIN_PATH, CFG_CHANGE_COMM_COMPS_PATH, CFG_CHANGE_HOSTED_FILENAME) _
                           )
        Select Case sReply
            Case CFG_CHANGE_ADDIN_PATH
                sPathCompMan = mBasic.SelectFolder( _
                               sTitle:="Select the folder for the AddIn instance of the CompManDev Workbook (escape to use the Application.UserLibraryPath)")
                If sPathCompMan = vbNullString Then
                    sPathCompMan = Application.UserLibraryPath ' Default because user escaped the selection
                Else
                    '~~ Assure trust in this location and save it to the CompMan.cfg file
                    mCfg.CompManAddinPath = sPathCompMan
                End If
            
            Case CFG_CHANGE_COMM_COMPS_PATH
                sPathCommon = mBasic.SelectFolder("Select the root folder for the ""Common Component Workbooks""")
                mCfg.CommonComponentsBasePath = sPathCommon

            Case CFG_CHANGE_HOSTED_FILENAME
                If mFl.SelectFile( _
                                    sel_init_path:=mCfg.CommonComponentsBasePath _
                                  , sel_title:="Select a file which - for example - is one indicating a Common Component(s) hosted in the coresponding Workbook." _
                                  , sel_result:=fl _
                                   ) _
                Then mCfg.HostedCommCompsFileName = fl.ShortName
        End Select
    Wend
    
    mDat.TrustThisFolder FolderPath:=mCfg.CompManAddinPath, TrustSubfolders:=False

xt: Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
        Case mErH.ErrMsgDefaultButton: End
    End Select
End Sub


Public Function HostedCommComps() As Collection
' -------------------------------------------------
' Returns a collection of file objects which
' specify hosted Common Components.
' -------------------------------------------------
            
    Dim cll             As New Collection
    Dim fso             As New FileSystemObject
    Dim fl              As FILE
    Dim v               As Variant
    Dim s               As String
    Dim ts              As TextStream
    Dim aNames()        As String
    Dim sParentFolder   As String
    Dim i               As Long
    Dim cllHosted       As New Collection
    Dim cllWbk          As New Collection
    
    mFl.Exists xst_file:=mCfg.CommonComponentsBasePath & "\" & mCfg.HostedCommCompsFileName & "*" _
               , xst_cll:=cll
    
    For Each v In cll
        Set fl = v
        sParentFolder = fl.ParentFolder
        mFl.Exists xst_file:=sParentFolder & "\*.xlsm", xst_cll:=cllWbk
'        Debug.Print cllWbk(1)
        
        Set ts = fl.OpenAsTextStream
        aNames = Split(ts.ReadAll, vbLf)
        Set ts = Nothing
        For i = LBound(aNames) To UBound(aNames)
            Debug.Print fl.ParentFolder & ": " & aNames(i)
        Next i
    Next v
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.name & ": mCfg." & sProc
End Function

