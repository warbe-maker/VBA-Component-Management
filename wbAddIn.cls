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
' ---------------------------------------------------------------------------
' Common Component Workbook CompManDev.xlsm
' - When setup this Workbook also exists as AddIn Workbook CompMan.xlma
' - When renewed, this Workbook is open as Addin Workbook CompMan.xlma
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
' W. Rauschenberger, Berlin Nov 2020
' See Github repo
' ---------------------------------------------------------------------------
Private Const ADDIN_WORKBOOK    As String = "CompMan.xlam"              ' Extension adjusted when the above is saved as addin
Private Const ADDIN_FORMAT      As Long = xlOpenXMLAddIn                ' = ... needs adjustment when the above is changed
Private Const ADDIN_VERSION     As String = "4.0"                       ' Allows to check the success of an Addin renew
Private Const DEVLP_WORKBOOK    As String = "CompManDev.xlsm"           ' Extension depends on Excel versions
Private Const DEVLP_FORMAT As Long = xlOpenXMLWorkbookMacroEnabled ' = ... needs adjustment when the above is changed

Public Property Get AddInInstanceFullName() As String:  AddInInstanceFullName = AddInPath & DBSLASH & AddInInstanceName:    End Property

Public Property Get AddInInstanceName() As String:      AddInInstanceName = ADDIN_WORKBOOK:                                 End Property

Private Property Get AddInPath() As String:             AddInPath = mCfg.CompManAddinPath:                           End Property

Public Property Get DevlpInstanceFullName() As String
Dim fso As New FileSystemObject
    DevlpInstanceFullName = mCfg.CommonComponentsBasePath & DBSLASH _
                          & fso.GetBaseName(DevlpInstanceName) & DBSLASH _
                          & DevlpInstanceName
End Property

Public Property Get DevlpInstanceName() As String:      DevlpInstanceName = DEVLP_WORKBOOK:                                 End Property

Public Property Get IsAddinInstance() As Boolean:       IsAddinInstance = ThisWorkbook.name = AddInInstanceName:            End Property

Public Property Get IsDevlpInstance() As Boolean:       IsDevlpInstance = ThisWorkbook.name = Me.DevlpInstanceName:         End Property

Public Property Get xlAddInFormat() As Long:            xlAddInFormat = ADDIN_FORMAT:                                       End Property

Public Property Get xlDevlpFormat() As Long:            xlDevlpFormat = DEVLP_FORMAT:                                       End Property

Public Function AddInVersion(Optional ByRef sVersion As String) As String
    sVersion = ADDIN_VERSION
    AddInVersion = sVersion
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "wbCompMan" & "." & sProc
End Function

Public Sub Renew()
    Me.Activate
    mAddIn.Renew ' The mAddin is declared private by intention and thus not available ti the menu bar
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If IsDevlpInstance Then mAddIn.ControlItemRenewRemove
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Const PROC      As String = "Workbook_BeforeSave"
    
    On Error GoTo eh
    Dim wb As Workbook

    If Not wbAddIn.IsDevlpInstance Then Exit Sub
    
    Set wb = ThisWorkbook
    mErH.BoP ErrSrc(PROC)
        
    Application.EnableEvents = False
    wb.Save
    Application.EnableEvents = True
    mCompMan.ExportChangedComponents wb ' includes management of updated used Common Components
    If Not wb.Saved Then
        Application.EnableEvents = False
        wb.Save        ' usually not required
        Application.EnableEvents = True
    End If
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Private Sub Workbook_Open()

    Dim fso As New FileSystemObject

    If IsDevlpInstance Then
        mAddIn.ControlItemRenewAdd
        mCfg.Asserted ' assert a valid CompMan basic configuratuion
        If mAddIn.AddInInstanceWorkbookIsOpen Then
            '~~ Essential! When the Development instance Workbook is opened, used Common Components
            '~~ the origin code has changed can only be updated when the Addin instance Workbook is
            '~~ open as well. However, since the VBProject has no reference to the Addin the Addin
            '~~ will not be opened when the Development instance is the only Workbook open.
            '~~ I.e. only when a Workbook referring to the Addin is opened the Addin will be opened as well.
            mAddIn.UpdateUsedCommComps
        End If
    End If
        
    With ThisWorkbook
        '~~ Make sure the VBProject name is the Workbook's base name
        '~~ Note that this is required to correctly identify (save, remove, restore) any reference
        '~~ to the Addin in open Workbook when Renew or SaveAsAddin is performed (see module mAddin)
        If .VBProject.name <> fso.GetBaseName(.name) Then .VBProject.name = fso.GetBaseName(.name)
        '~~ Unporotecting (the Addin) allows directly modify code in it - and saving it subsequently
        '~~ as new Development instance (see method SaveasDev in module mAddin)
        .Unprotect
    End With

End Sub

