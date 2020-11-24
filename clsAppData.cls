VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' ------------------------------------------------------------------------
' Class Module clsAppData
'       Provides a set of methods, functions, and procedures to write and
'       read so called "Session Persistent Properties" as named values
'       either in a file ("Private Profile") or the registry.
'       While the registry allows a hierarchy of keys, the values
'       in a file are organized by [sections] which makes it slightly
'       more complex to organize hierarchies.
'       The properties' usage differs as follows:
'
'                       Is represented  Is Represented in
'           Property    in a File as    the Registry as
'           ----------  --------------  ------------------
'           Subject     Path/Filename   SubKey under HKEY....
'           Extension   File-Extension  n/a
'           Aspect      Section         SubKey under Subject SubKey
'           HKey        n/a             HKEY_...
'
' Properties: In logical order!
'           - Location      Specifies where the named values are to be stored
'           - Subject       The subject for the maintained "Sesion Persistent Properties"
'                           may be an Excel Workbook or any other Office Application's object
'           - Extension     The file name extension, aplicable only when Location=File
'           - Aspect        Groups "Session Persistent Properties", defaults to "Common" when
'                           not used/provided
'           - HKey          The Registry "hive key" for the Subject subkey (applicable only
'                           with Location=spp_Registry
' Public Methods:
'          - AspectRemove   Removes a section including all its keys. Aspect defaults to the
'                           Property 'Aspect' when not provided as parameter.
'          - Aspects        Dictionary of all Aspects
'          - AspectsDisplay Displays all Aspectes under a given Subject
'          - DefaultLet     Stores default value for a certain value name which us returned
'                           in case ValueGet is used with a value name which had never been
'                           writte/saved.
'          - Exists         Returns TRUE when the specified Subject, Aspect, or Name exists
'                           at the given Location. Uses the Properties .Location, .Subject,
'                           .Aspect, and .HKey as defaults for not provided parameters.
'          - NameRemove     Removes a name entry from the provided Aspect
'                           .Aspect and .Name when not provided as parameters
'          - NamesDisplay   Displays all properties and their value (for test purpose only!)
'          - AspectRemove   Removes a section including all its keys. Aspect defaults to the
'                           Property 'Aspect' when not provided as parameter
'          - ValuesDisplay  Displays all values of a given Subject and Aspect at a given Location.
'                           Uses the Properties .Location, .Subject, .Aspect, and .HKey as
'                           defaults for not provided parameters.
'          - ValueGet       Returns the value of a named value. Uses the Properties .Location,
'                           .Subject, .Aspect, .HKey as defaults for not provided parameters.
'                           When a variable for the target value's data type is provided, the
'                           return string is converted to it.
'          - ValueLet       Saves a value under a given name. Uses the Properties .Location,
'                          .Subject, .Aspect, .HKey as defaults for not provided parameters.
'          - SubjectRemove Removes the given Subject (the file representing it or the subkey
'                          under the HKey. Uses the Property .Location for the not provided
'                          parameter.
'
' Author W. Rauschenberger, Berlin, Sept 2019
' ------------------------------------------------------------------------------------------------
Private Const sMod = "clsAppData"

Private Declare PtrSafe Function WritePrivateProfileString _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpString As Any, _
                ByVal lpName As String) As Long

Private Declare PtrSafe Function GetPrivateProfileString _
                Lib "kernel32" Alias "GetPrivateProfileStringA" _
               (ByVal lpApplicationName As String, _
                ByVal lpKeyName As Any, _
                ByVal lpDefault As String, _
                ByVal lpReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpName As String) As Long
                             
Private Declare PtrSafe Function DeletePrivateProfileSection _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal section As String, _
                ByVal NoKey As Long, _
                ByVal NoSetting As Long, _
                ByVal Name As String) As Long

Private Declare PtrSafe Function DeletePrivateProfileKey _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal section As String, _
                ByVal Key As String, _
                ByVal Setting As Long, _
                ByVal Name As String) As Long
                 
Private Declare PtrSafe Function GetPrivateProfileSectionNames _
                Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" _
               (ByVal lpszReturnBuffer As String, _
                ByVal nSize As Long, _
                ByVal lpName As String) As Long
                 
Private Declare PtrSafe Function GetPrivateProfileSection _
                Lib "kernel32" Alias "GetPrivateProfileSectionA" _
               (ByVal section As String, _
                ByVal Buffer As String, _
                ByVal Size As Long, _
                ByVal Name As String) As Long

Private Const HKEY_CLASSES_ROOT         As Long = &H80000000
Private Const HKEY_CURRENT_CONFIG       As Long = &H80000005
Private Const HKEY_CURRENT_USER         As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Private Const HKEY_USERS                As Long = &H80000003
Private Const REG_MAX_VALUE_LENGTH      As Long = &H100

' -------------------------------------------------------------
' The 'private' properties of each "Sesion Persistent Property"
' -------------------------------------------------------------
Private lHKey       As enHKey       ' The current effective HKEY (Defaults to spp_HKEY_CURRENT_USER
Private lLocation   As enLocation   ' The current effective Location (defaults to spp_Registry)
Private sAspect     As String       ' The current effective Aspect (defaults to "Common")
Private sExtension  As String       ' The "Private Profile" file's extension (defaults to .ini)
Private sSubject    As String       ' The current effective Subject (name of the "Private Profile" file or the SubKey under the HKey)
Private dctDefaults As Dictionary   ' Default values for value names

Private Sub Class_Initialize()
' ----------------------------------------------------
' Proper default values are essential for all methods.
' These defaults or any values provided through the
' corresponding properties are used throughout this
' class module.
' ----------------------------------------------------
    '~~> Assign defaults
    lLocation = spp_Registry
    lHKey = spp_HKEY_CURRENT_USER
    sExtension = ".ini"
    sAspect = COMMON_
    Set dctDefaults = New Dictionary    ' By default there are no default values
    
End Sub

Private Sub Class_Terminate()
    Set dctDefaults = Nothing
End Sub

Friend Property Get Aspect() As String:             Aspect = sAspect:                           End Property

Friend Property Let Aspect(s As String)            ' The Current Aspect's name
    sAspect = s:    If lLocation = spp_Registry Then AspectAdjustToRegistry sAspect
End Property

Friend Property Let Extension(s As String)         ' File extension when Location = spp_File
    If left$(s, 1) = "." Then sExtension = s Else sExtension = "." & s ' Prefix extension with a .
    SubjectAdjustExtension
End Property

Friend Property Get HKey() As enHKey:               HKey = lHKey:                               End Property

Friend Property Let HKey(l As enHKey):              lHKey = l:                                  End Property

Private Property Get HKeyHex(Optional ByVal lHKey As enHKey = 0) As Long

    '~~ Use the default values for not provided parameters
    If lHKey = 0 Then lHKey = Me.HKey
    Select Case lHKey
        Case spp_HKEY_CLASSES_ROOT:    HKeyHex = HKEY_CLASSES_ROOT
        Case spp_HKEY_CURRENT_CONFIG:  HKeyHex = HKEY_CURRENT_CONFIG
        Case spp_HKEY_CURRENT_USER:    HKeyHex = HKEY_CURRENT_USER
        Case spp_HKEY_LOCAL_MACHINE:   HKeyHex = HKEY_LOCAL_MACHINE
        Case spp_HKEY_USERS:           HKeyHex = HKEY_USERS
        Case Else:                     HKeyHex = HKEY_CURRENT_USER
    End Select

End Property

Friend Property Get Location() As enLocation:       Location = lLocation:                       End Property

Public Property Let Location(l As enLocation)            ' Determines the location of the maintained properties
    lLocation = l
    If lLocation = spp_Registry Then
        SubjectAdjustToRegistry
        AspectAdjustToRegistry Aspect
    End If
End Property

Public Property Get MaxRegValueLength() As Long:    MaxRegValueLength = REG_MAX_VALUE_LENGTH:   End Property

Friend Property Get Subject() As String:            Subject = sSubject:                         End Property

Public Property Let Subject(s As String)
    sSubject = s:   If lLocation = spp_Registry Then SubjectAdjustToRegistry Else SubjectAdjustExtension
End Property

Private Sub AspectAdjustToRegistry(ByRef sAspect As String)
    If Len(sAspect) > REG_MAX_VALUE_LENGTH Then
        sAspect = "..." & Right$(sAspect, REG_MAX_VALUE_LENGTH - 3)
    End If
End Sub

Public Sub AspectRemove(Optional ByVal lLocation As enLocation = 0, _
                        Optional ByVal sSubject As String = vbNullString, _
                        Optional ByVal sAspect As String = vbNullString, _
                        Optional ByVal lHKey As enHKey = 0)
' ------------------------------------------------------------------------
' Removes all properties of 'sAspect'.
' lLocation=spp_File:     The 'sAspect' is represented as a section in the
'                         file named 'Subject'
' lLocation=spp_Registry: The 'sAspect' is represented as a subkey in the
'                         registry path 'HKey'\'Subject'\
' ------------------------------------------------------------------------
Const PROC  As String = "AspectRemove"
Dim wss     As IWshShell3
Dim sPath   As String

    On Error GoTo eh
    
    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If sAspect = vbNullString Then sAspect = Me.Aspect
    If lHKey = 0 Then lHKey = Me.HKey
    
    
    Select Case lLocation
        Case spp_File
            If Not Exists(spp_Subjct, lLocation, sSubject, , , lHKey) Then Exit Sub
            If Not Exists(spp_Aspect, lLocation, sSubject, sAspect, , lHKey) Then Exit Sub
            DeletePrivateProfileSection sAspect, 0, 0, sSubject
            
        Case spp_Registry
            sPath = HKeyToStrng(lHKey) & "\" & sSubject & "\" & sAspect & "\"
            Set wss = CreateObject("WScript.Shell")
            On Error Resume Next
            wss.RegDelete sPath
            Set wss = Nothing
    End Select
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Function Aspects(Optional ByVal lLocation As enLocation = 0, _
                        Optional ByVal sSubject As String = vbNullString, _
                        Optional ByVal lHKey As enHKey = 0) As Dictionary
' -------------------------------------------------------------------------
' Returns a Dictionary of all Aspects.
' -------------------------------------------------------------------------
Const PROC = "Aspects"
Dim asAspects()     As String
Dim avKeys()        As Variant
Dim dct             As Dictionary
Dim i               As Long
Dim iLen            As Long
Dim obj             As Object
Dim strBuffer       As String
Dim v               As Variant
Dim fso             As File

    On Error GoTo eh
    Set dct = New Dictionary
    Set Aspects = dct
    
    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If lHKey = 0 Then lHKey = Me.HKey
            
    Select Case lLocation
        Case spp_File
            If Not Exists(spp_Subjct, lLocation, sSubject, , , lHKey) _
            Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The provided subject '" & sSubject & "' does not exist!"
            With New FileSystemObject
                If .GetFile(sSubject).Size = 0 Then Exit Function
            End With
            Do While (iLen = Len(strBuffer) - 2) Or (iLen = 0)
                If strBuffer = vbNullString _
                Then strBuffer = mBasic.Space$(256) _
                Else strBuffer = String(Len(strBuffer) * 2, 0)
                iLen = GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), sSubject)
            Loop
            strBuffer = left$(strBuffer, iLen)
            
            If Len(strBuffer) <> 0 Then
                i = 0
                asAspects = Split(strBuffer, vbNullChar)
                For i = LBound(asAspects) To UBound(asAspects)
                    If Len(asAspects(i)) <> 0 Then
                        mBasic.DictAdd dct, asAspects(i), asAspects(i), dct_ascending
                    End If
                Next i
            End If
    
        Case spp_Registry
            On Error Resume Next
            Set obj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\default:StdRegProv")
            If Err.Number <> 0 Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "There where no aspectes at all found in the registry!"
            obj.EnumKey HKeyHex(lHKey), sSubject & "\", avKeys
            For Each v In avKeys
                '~~ An empty Aspect means no Aspects
                If v <> vbNullString Then
                    mBasic.DictAdd dct, v, CStr(v), dct_ascending
                End If
            Next v
            
    End Select
    Set Aspects = dct
    Exit Function
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Sub AspectsDisplay(Optional ByVal lLocation As enLocation = 0, _
                          Optional ByVal sSubject As String = vbNullString, _
                          Optional ByVal lHKey As enHKey = 0)
' ---------------------------------------------------------------------------
'
' ---------------------------------------------------------------------------
Const PROC          As String = "DisplayAspects"
Dim sAspects        As String
Dim dct             As Dictionary
Dim v               As Variant

    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If lHKey = 0 Then lHKey = Me.HKey
    
    If Not Exists(spp_Subjct, lLocation, sSubject, , , lHKey) _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The provided subject '" & sSubject & "' does not exist!"
    
    Set dct = Aspects(lLocation, sSubject, lHKey)
    sAspects = dct.Count & " Aspect(s) of subject" & vbLf & _
               "'" & Split(sSubject, "\")(UBound(Split(sSubject, "\"))) & "':" & vbLf & vbLf
    For Each v In dct
        sAspects = sAspects & "- '" & v & "'" & vbLf
    Next v
    MsgBox sAspects, vbOKOnly, "Info about Session Persistent Properties in '" & LocationToString(lLocation) & "'"
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Private Function DefaultGet(ByVal sName As String, _
                   Optional ByRef vTarget As Variant, _
                   Optional ByVal sAspect As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Returns the default value for a value named (sName) when it exists in the
' dictionary of Default values (dctDefaults). When no Aspect (sAspect) is
' provided the default available for all Aspects is returned in case one
' exists.
' If (vTarget) is provided and no default value exists the default for the
' data type of vTarget is returned (Boolean=False, String=vbNullString,
' Numeric=0).
' When an Aspect is provided the Aspect-specific default value is returned
' when existing.
' Note: The function is public for testing it may be private in production.
' ----------------------------------------------------------------------------
Const PROC      As String = "DefaultGet"
Dim sKey        As String
Dim vValue      As Variant

    On Error GoTo eh
    
    If Not dctDefaults Is Nothing Then
        If sAspect <> vbNullString Then
            '~~ Try to return the default value for the Aspect/Name
            sKey = sAspect & "\" & sName
            If dctDefaults.Exists(sKey) Then
                DefaultGet = dctDefaults.Item(sKey)
                Exit Function
            End If
        Else
            '~~ Try to return the default value for the Name regardless of the Aspect
            sKey = vbNullString & "\" & sName
            If dctDefaults.Exists(sKey) Then
                DefaultGet = dctDefaults.Item(sKey)
                Exit Function
            End If
        End If
    End If
    
    '~~ No default value found or no defaults available at all
    Select Case VarType(vTarget)
        Case vbBoolean: vTarget = False:    DefaultGet = vTarget
        Case vbDate:    vTarget = Now():    DefaultGet = vTarget
        Case vbByte, _
             vbCurrency, _
             vbDecimal, _
             vbDouble, _
             vbInteger, _
             vbLong, _
             vbSingle: vTarget = 0: DefaultGet = vTarget
#If Win64 Then
        Case vbLongLong
            vTarget = 0
            DefaultGet = vTarget
#End If
        Case Else:                  DefaultGet = vbNullString   ' No vTarget provided
    End Select
    Exit Function
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Sub DefaultLet(ByVal sName As String, _
                      ByVal vValue As Variant, _
             Optional ByVal sAspect As String = vbNullString)
' -----------------------------------------------------------
' Stores the default (vValue) for the value named (sName) of
' the aspect (sAspect) in the Dictionary dctDefaults. When
' Aspect (sAspet) is not provided the saved default wil be
' effective for all Aspects.
' The stored defaults comes into effect when the value named
' (sName) is retrieved before a value with that name had been
' saved.
' ----------------------------------------------------------
Dim sKey As String

    If dctDefaults Is Nothing Then Set dctDefaults = New Dictionary
    
    sKey = sAspect & "\" & sName
    If dctDefaults.Exists(sKey) Then
        dctDefaults.Item(sKey) = vValue
    Else
        mBasic.DictAdd dctDefaults, sKey, vValue, dct_ascending
    End If

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mVBP" & "." & sProc
End Function

Public Function Exists(Optional ByVal sppItem As enItem = 0, _
                       Optional ByVal lLocation As enLocation = 0, _
                       Optional ByVal sSubject As String = vbNullString, _
                       Optional ByVal sAspect As String = vbNullString, _
                       Optional ByVal sName As String = vbNullString, _
                       Optional ByVal lHKey As enHKey = 0) As Boolean
' ------------------------------------------------------------------------
' - sppItem = spp_Subjct: Returns TRUE when the subject (sSubject) exists
'   - when lLocation = spp_File as file or
'   - when lLocation = spp_Registry as a subkey under the HKey (lHKey)
' - sppItem = spp_Aspect: Returne TRUE when the aspect (sAspect) exists
'   under the given location under the given subject. Ends with an error
'   when the subject (sSubject) does not exist.
' - sppItem = spp_VaName: Returns TRUE when the name exists under the given
'   aspect (sAspect) under the given subject (sSubject) at the given
'   location (lLocation). Ends with an error when the subject (sSubject)
'   or the aspect (sAspect) do not exist.
' Note: The function is public to support DebugAssert while testing.
' ------------------------------------------------------------------------
Const PROC  As String = "Exists"
Dim obj     As Object
Dim sPath   As String
Dim avKeys  As Variant
Dim v       As Variant
    
    On Error GoTo eh

    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If sAspect = vbNullString Then sAspect = Me.Aspect
    If lHKey = 0 Then lHKey = Me.HKey
    
'    '~~ Values validation
'    If lLocation = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "There is no Location provided and there is no default Location either!"
'    If sSubject = vbNullString Then Err.Raise AppErr(2), ErrSrc(PROC), "There is no Subject provided and there is no default Socation either!"
'    If sAspect = vbNullString Then Err.Raise AppErr(3), ErrSrc(PROC), "There is no Aspect provided and there is no default Aspect either!"
'    If lHKey = 0 Then Err.Raise AppErr(4), ErrSrc(PROC), "There is no HKey provided and there is no default HKey either!"
    
    Exists = False
    Select Case sppItem
        
        Case spp_Subjct
            Select Case lLocation
                Case spp_File
                    With New Scripting.FileSystemObject
                        Exists = .FileExists(sSubject)
                    End With
                
                Case spp_Registry
                    lHKey = lHKey:  If lHKey = 0 Then lHKey = Me.HKey
                    Set obj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\default:StdRegProv")
                    sPath = sSubject & "\"
                    sPath = Replace(sPath, "\\", "\")
                    On Error Resume Next
                    obj.EnumKey HKeyHex(lHKey), sPath, avKeys
                    If Err.Number <> 0 Then Exit Function ' Subject does not exist
                    For Each v In avKeys
                        If v <> vbNullString Then
                            Exists = True
                            Exit For
                        End If
                    Next v
            End Select
                
        Case spp_Aspect
            '~~ The distintion between the Locations is done in the Aspects function
            lHKey = lHKey: If lHKey = 0 Then lHKey = Me.HKey
            '~~ Remove any [ ] from the Parameter sAspect
            sAspect = Replace(Replace(sAspect, "[", vbNullString), "]", vbNullString)
            Exists = Aspects(lLocation, sSubject, lHKey).Exists(sAspect)
        
        Case spp_VaName
            lHKey = lHKey: If lHKey = 0 Then lHKey = Me.HKey
            If Not Exists(spp_Subjct, lLocation, sSubject, , , lHKey) _
            Then Err.Raise mErH.AppErr(5), ErrSrc(PROC), "The subject '" & sSubject & "' does not exist!"
            If Not Exists(spp_Aspect, lLocation, sSubject, sAspect, , lHKey) _
            Then Err.Raise mErH.AppErr(6), ErrSrc(PROC), "The aspect '" & sAspect & "' does not exist under the given subject!"
            
            Exists = Names(lLocation, sSubject, sAspect, lHKey).Exists(sName)
    End Select
    Exit Function
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Function HKeyToStrng(Optional ByVal lHKey As enHKey = 0) As String
' -----------------------------------------------------------------------

    '~~ Use the default values for not provided parameters
    If lHKey = 0 Then lHKey = Me.HKey
    
    Select Case lHKey
        Case spp_HKEY_CLASSES_ROOT:     HKeyToStrng = "HKEY_CLASSES_ROOT"
        Case spp_HKEY_CURRENT_CONFIG:   HKeyToStrng = "HKEY_CURRENT_CONFIG"
        Case spp_HKEY_CURRENT_USER:     HKeyToStrng = "HKEY_CURRENT_USER"
        Case spp_HKEY_LOCAL_MACHINE:    HKeyToStrng = "HKEY_LOCAL_MACHINE"
        Case spp_HKEY_USERS:            HKeyToStrng = "HKEY_USERS"
    End Select
    
End Function

Private Function LocationToString(Optional ByVal lLocation As enLocation = 0) As String

    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    Select Case lLocation
        Case spp_File:      LocationToString = "File"
        Case spp_Registry:  LocationToString = "Registry"
    End Select
    
End Function

Public Sub NameRemove(ByVal sName As String, _
             Optional ByVal sAspect As String = vbNullString, _
             Optional ByVal sSubject As String = vbNullString, _
             Optional ByVal lLocation As enLocation, _
             Optional ByVal lHKey As enHKey = 0)
' --------------------------------------------------------------
' Removes the name (sName) from the Aspect (sAspect). For any
' optional parameter not provided the corresponding class
' property value is used - which may still be the default).
' --------------------------------------------------------------
Const PROC  As String = "NameRemove"
Dim sPath   As String
Dim wss     As IWshShell3

    On Error GoTo eh
    
    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If sAspect = vbNullString Then sAspect = Me.Aspect
    If lHKey = 0 Then lHKey = Me.HKey
    
    '~~ Value validation
    If Not Exists(spp_Subjct, lLocation, sSubject, , , lHKey) _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The subject '" & sSubject & "' does not exist!"
    If Not Exists(spp_Aspect, lLocation, sSubject, sAspect, , lHKey) _
    Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "The aspect '" & sAspect & "' does not exist!"
    
    If Not Exists(spp_VaName, lLocation, sSubject, sAspect, sName, lHKey) Then Exit Sub
    
    Select Case lLocation
        Case spp_File
            DeletePrivateProfileKey sAspect, sName, 0, sSubject
    
        Case spp_Registry
            sPath = HKeyToStrng & "\" & sSubject & "\" & sAspect & "\" & sName
            Set wss = CreateObject("WScript.Shell")
            On Error Resume Next
            wss.RegDelete sPath
            Set wss = Nothing
            If Exists(spp_VaName, spp_Registry, sSubject, sAspect, sName, lHKey) _
            Then Err.Raise mErH.AppErr(3), ErrSrc(PROC), "The name '" & sName & "' had not been removed!"
    End Select

xt: Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Private Function Names(Optional lLocation As enLocation = 0, _
                       Optional sSubject As String = vbNullString, _
                       Optional sAspect As String = vbNullString, _
                       Optional lHKey As enHKey = 0) As Dictionary
' -----------------------------------------------------------------
' Returns a Dictionary with all names under the aspect (sAspect).
' When no aspect is provided, all names under all aspects of the
' subject (sSubject) are returned. The name becomes the key in
' order to primarily support the existence check. The concerned
' aspect is the the item. In case a name is used under several
' Aspects, the item lists them all delimited by a comma.
' -----------------------------------------------------------------
Const PROC          As String = "Names"
Dim asNames()       As String
Dim avNames()       As Variant
Dim avTypes()       As Variant
Dim dctAspects      As Dictionary
Dim dctNames        As Dictionary
Dim i               As Long
Dim lResult         As Long
Dim obj             As Object
Dim sNames          As String
Dim sPath           As String
Dim strBuffer       As String
Dim sValue          As String
Dim v               As Variant
Dim sItem           As String
Dim vAspect         As Variant

    Set dctNames = New Dictionary
    Set dctAspects = New Dictionary
    Set Names = dctNames    ' Empty in case no names are returned
    
    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If sAspect = vbNullString Then sAspect = Me.Aspect
    If lHKey = 0 Then lHKey = Me.HKey
    
    '~~ Values validation
    If lLocation = 0 Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "There is no location provided and there is no default location either!"
    
    If lLocation = spp_Registry Then
        If lHKey = 0 Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "There is no HKey provided and there is no default HKey available either!"
    End If
    
    If sSubject = vbNullString Then sSubject = Me.Subject Else sSubject = sSubject
    If sSubject = vbNullString Then Err.Raise mErH.AppErr(3), ErrSrc(PROC), "There is no subject provided and no default value available either!"
     
    Select Case lLocation
        Case spp_File
            If sAspect <> vbNullString Then
                '~~> Retrieve the names for the provided Aspect
                strBuffer = mBasic.Space$(32767)
                lResult = GetPrivateProfileString(sAspect, vbNullString, vbNullString, strBuffer, Len(strBuffer), sSubject)
                sNames = left$(strBuffer, lResult)
            
                If sNames <> vbNullString Then                                         ' If there were any names
                    asNames = Split(sNames, vbNullChar)                      ' have them split into an array
                    For i = LBound(asNames) To UBound(asNames)
                        If Len(asNames(i)) <> 0 Then
                            mBasic.DictAdd dctNames, sAspect, asNames(i), dct_ascending
                        End If
                    Next i
                End If
            Else
                '~~> Retrieve the names of all Aspects
                Set dctAspects = Aspects(lLocation, sSubject, lHKey)
                For Each v In dctAspects
                    sAspect = v
                    strBuffer = mBasic.Space$(32767)
                    lResult = GetPrivateProfileString(sAspect, vbNullString, vbNullString, strBuffer, Len(strBuffer), sSubject)
                    sNames = left$(strBuffer, lResult)
                
                    If sNames <> vbNullString Then                                         ' If there were any names
                        asNames = Split(sNames, vbNullChar)                      ' have them split into an array
                        For i = LBound(asNames) To UBound(asNames)
                            If Len(asNames(i)) <> 0 Then
                                If dctNames.Exists(asNames(i)) Then
                                    sItem = dctNames.Item(asNames(i))
                                    dctNames.Item(asNames(i)) = sItem & "," & sAspect
                                Else
                                    mBasic.DictAdd dctNames, asNames(i), sAspect & "/" & asNames(i), dct_ascending
                                End If
                            End If
                        Next i
                    End If
                    
                Next v
            End If
        
        Case spp_Registry
            Set obj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & "." & "\root\default:StdRegProv")
            If sAspect <> vbNullString Then
                sPath = sSubject & "\" & sAspect & "\"
                On Error Resume Next
                obj.EnumValues HKeyHex(lHKey), sPath, avNames, avTypes
                If Err.Number <> 0 Then Exit Function
                For Each v In avNames
                    '~~ An empty Name means no Names
                    If v <> vbNullString Then
                        mBasic.DictAdd dctNames, v, CStr(v), dct_ascending
                    End If
                Next v
            Else
                For Each vAspect In dctAspects
                    sPath = sSubject & "\" & vAspect & "\"
                    On Error Resume Next
                    obj.EnumValues HKeyHex(lHKey), sPath, avNames, avTypes
                    If Err.Number <> 0 Then Exit Function
                    For Each v In avNames
                        If v <> vbNullString Then
                            mBasic.DictAdd dctNames, v, vAspect & "/" & CStr(v), dct_ascending
                        End If
                    Next v
                Next vAspect
            End If
            Set obj = Nothing
    End Select
    Set Names = dctNames

xt: Exit Function
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Sub NamesDisplay(Optional ByVal lLocation As enLocation = 0, _
                        Optional ByVal sSubject As String = vbNullString, _
                        Optional ByVal sAspect As String = vbNullString, _
                        Optional ByVal lHKey As enHKey = 0)
' ---------------------------------------------------------------------------
'
' ---------------------------------------------------------------------------
Const PROC  As String = "NamesDisplay"
Dim sNames  As String
Dim dct     As Dictionary
Dim v       As Variant

    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If lHKey = 0 Then lHKey = Me.HKey
    If sAspect = vbNullString Then sAspect = Me.Aspect
    
    If Not Exists(spp_Subjct, lLocation, sSubject, , , lHKey) _
    Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The provided subject '" & sSubject & "' does not exist!"
    
    Set dct = Names(lLocation, sSubject, sAspect, lHKey)
    sNames = dct.Count & " Name(s) of subject" & vbLf & _
               "'" & sSubject & "':" & vbLf & vbLf
    For Each v In dct
        sNames = sNames & "- '" & v & "'" & vbLf
    Next v
    MsgBox sNames, vbOKOnly, "Info about Session Persistent Properties in '" & LocationToString(lLocation) & "'"
    
xt: Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Private Function RegKeyRead(ByVal sRegKey As String) As String
' ----------------------------------------------------
'
' ----------------------------------------------------
Dim wss As IWshShell3

  On Error Resume Next
  
  RegKeyRead = vbNullString
  Set wss = CreateObject("WScript.Shell")
  RegKeyRead = wss.RegRead(sRegKey)
  Set wss = Nothing

End Function

Private Function RegTypeStrng(ByVal lRegType As enRegDataType) As String
    Select Case lRegType
        Case spp_REG_BINARY:    RegTypeStrng = "REG_BINARY"
        Case spp_REG_DWORD:     RegTypeStrng = "REG_DWORD"
        Case spp_REG_EXPAND_SZ: RegTypeStrng = "REG_EXPAND_SZ"
        Case spp_REG_SZ:        RegTypeStrng = "REG_SZ"
        Case Else:              RegTypeStrng = "REG_SZ"    ' Other types not supported by RegWrite
    End Select
End Function

Private Sub SubjectAdjustExtension()
' -------------------------------------
' Adjust the extension when a new
' extension or subject is specified.
' -------------------------------------
Dim l As Long

    If InStr(1, sSubject, "\") <> 0 Then
        '~~> When the extension is defined 'Subject' is a file path the extension is adjusted accordingly
        l = InStrRev(sSubject, ".")
        If l <> 0 Then
            '~~> Unstrip the existing extension and amend it with the new specified
            sSubject = left$(sSubject, l - 1) & sExtension
        Else
            sSubject = sSubject & sExtension
        End If
    End If

End Sub

Private Sub SubjectAdjustToRegistry()
    '~~> Make the Subject a valid registry key
    sSubject = Replace(sSubject, "\", "/")
    If Len(sSubject) > MaxRegValueLength Then
        sSubject = "..." & Right$(sSubject, MaxRegValueLength - 3)
    End If
End Sub

Public Sub SubjectRemove(Optional ByVal sSubject As String = vbNullString, _
                         Optional ByVal lLocation As enLocation = 0, _
                         Optional ByVal lHKey As enHKey = 0)
' --------------------------------------------------------------------------
' Returns TRUE when the Subject exists at the 'lLocation'.
' lLocation=spp_File:     The Subject is represented as a path/filename
' lLocation=spp_Registry: The Subject is represented as a subkeys
'                         of the 'HKey'
' --------------------------------------------------------------------------
Const PROC      As String = "SubjectRemove"
Dim sPath       As String
Dim dctAspects  As Dictionary
Dim v           As Variant
Dim wss         As IWshShell3

    On Error GoTo eh
    
    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If lHKey = 0 Then lHKey = Me.HKey

    If lLocation = 0 Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "There is no location specified from which a subject is to be removed!"
    If lLocation = spp_Registry Then
        If lHKey = 0 Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "There is no HKey specified from which a subject is to be removed!"
    End If
    If sSubject = vbNullString Then Err.Raise mErH.AppErr(3), ErrSrc(PROC), "There is not subject to be removed specified!"
    
    Select Case lLocation
        Case spp_File
            With New Scripting.FileSystemObject
                If .FileExists(sSubject) Then .DeleteFile sSubject
            End With
        
        Case spp_Registry
            Set wss = CreateObject("WScript.Shell")
            With wss
                '~~> At first we have to remove all the Subject's Aspects (subkeys)
                Set dctAspects = Aspects(spp_Registry, sSubject, lHKey)
                For Each v In dctAspects
                    sPath = HKeyToStrng(lHKey) & "\" & sSubject & "\" & v & "\"
                    On Error Resume Next
                    .RegDelete sPath
                    If Exists(spp_Aspect, spp_Registry, sSubject, v, , lHKey) _
                    Then Err.Raise mErH.AppErr(4), ErrSrc(PROC), "Aspect '" & v & "' (path '" & sPath & "') not deleted!"
                Next
                '~~> Finally remove the Subject
                sPath = HKeyToStrng & "\" & sSubject & "\"
                On Error Resume Next
                .RegDelete sPath
            End With
            Set wss = Nothing
            If Exists(spp_Subjct, lLocation, sSubject, , , lHKey) _
            Then Err.Raise mErH.AppErr(5), ErrSrc(PROC), "The subject '" & sSubject & "' (path '" & sPath & "') had not been deleted!"
    End Select
    
xt: Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Function ValueGet(ByVal sName As String, _
                Optional ByRef vTarget As Variant, _
                Optional ByVal lLocation As enLocation = 0, _
                Optional ByVal sSubject As String = vbNullString, _
                Optional ByVal sAspect As String = vbNullString, _
                Optional ByVal lHKey As enHKey = 0) As Variant
' -----------------------------------------------------------------
' Returns the value named (sName) either from the "Private Profile"
' file or the Registry depending on the current class property
' Location or the provided (lLocation). When a values named (sName)
' cannot be found a possibly specified default value (see method
' DefaultLet) is returned else a vbNullString. When the target
' value is provided by (vTarget) the value is returned in it and
' converted to the lTarget's data format. The default return value
' when the named vlue cannot be found is: Boolean=FALSE, String=
' vbNullstring, Numeric=0. Since all values are saved and retrieved
' as strings, the conversion to the original data type has to be
' done by the caller when no vTarget is provided.
' -----------------------------------------------------------------
Const PROC  As String = "ValueGet"
Dim lResult As Long
Dim sRegKey As String
Dim sRetVal As String
Dim vValue  As Variant
Dim wss     As IWshShell3

    On Error GoTo eh
    
    ValueGet = vbNullString
    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If sAspect = vbNullString Then sAspect = Me.Aspect
    If lHKey = 0 Then lHKey = Me.HKey
    
    '~~ Validate provided values
    If lLocation = 0 Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "No location provided and there is no default location available either!"
    If sSubject = vbNullString Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "No subject provided and there is no default subject available either!"
    If sAspect = vbNullString Then Err.Raise mErH.AppErr(3), ErrSrc(PROC), "No aspect provided and there is no default aspect available either!   "
    If lLocation = spp_Registry Then
        If lHKey = 0 Then Err.Raise mErH.AppErr(4), ErrSrc(PROC), "No HKey provided and there is no default HKey available either!"
    End If
    
    '~~ If the either subject, aspect, or name doesn't exist a default value is returned when specified
    Select Case lLocation
        Case spp_File
            sRetVal = String(32767, 0)
            '~~> Execute the call and get the string lenght
            lResult = GetPrivateProfileString(sAspect, sName, vbNullString, sRetVal, Len(sRetVal), sSubject)
            vValue = left$(sRetVal, lResult)
    
        Case spp_Registry
            If Exists(spp_Aspect, lLocation, sSubject, sAspect, , lHKey) Then
                If Exists(spp_VaName, lLocation, sSubject, sAspect, sName, lHKey) Then
                    sRegKey = HKeyToStrng & "\" & sSubject & "\" & sAspect & "\" & sName
                    Set wss = CreateObject("WScript.Shell")
                    On Error Resume Next
                    vValue = wss.RegRead(sRegKey)
                    Set wss = Nothing
                End If
            End If
    End Select
    
    If Len(vValue) = 0 Then vValue = DefaultGet(sName, vTarget, sAspect)
    
    Select Case VarType(vTarget)
        Case vbBoolean:     vTarget = VBA.CBool(vValue):    ValueGet = vTarget
        Case vbByte:        vTarget = VBA.CByte(vValue):    ValueGet = vTarget
        Case vbCurrency:    vTarget = VBA.CCur(vValue):     ValueGet = vTarget
        Case vbDate:        vTarget = VBA.CDate(vValue):    ValueGet = vTarget
        Case vbDecimal:     vTarget = VBA.CDec(vValue):     ValueGet = vTarget
        Case vbDouble:      vTarget = VBA.CDbl(vValue):     ValueGet = vTarget
        Case vbInteger:     vTarget = VBA.CInt(vValue):     ValueGet = vTarget
        Case vbLong:        vTarget = VBA.CLng(vValue):     ValueGet = vTarget
#If Win64 Then
'        Case vbLongLong:    vTarget = VBA.CLngLng(vValue):  ValueGet = vTarget
#End If
        Case vbSingle:      vTarget = VBA.CSng(vValue):     ValueGet = vTarget
        Case vbString:      vTarget = VBA.CStr(vValue):     ValueGet = vTarget
        Case Else:                                          ValueGet = vValue       ' No vTarget provided
    End Select
    
xt: Exit Function
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Sub ValueLet(ByVal sName As String, _
                    ByVal vValue As Variant, _
           Optional ByVal lType As enRegDataType = 99, _
           Optional ByVal sAspect As String = vbNullString, _
           Optional ByVal sSubject As String = vbNullString, _
           Optional ByVal lLocation As enLocation = 0, _
           Optional ByVal lHKey As enHKey = 0)
' ------------------------------------------------------------
' Writes the value (vValue) under the name (sName) under the
' provided aspect (sAspect). Any name, aspect or subject not
' existing is created on the fly, else the value is just
' updated.
' - lLocation = spp_File stores aspect, name and value in the
'   file named sSubject in the form:
'   [aspect]
'   name=value
'
' - lLocation = spp_Registry stores subject, aspect, name and
'   value in the form:
'   HKey\subject\aspect name value
'   whereby HKey = HKEY_CURRENT_USER
' -------------------------------------------------------------
Const PROC      As String = "ValueLet"
Dim lChars      As Long
Dim sPath       As String
Dim wss         As IWshShell3
Dim sRegType    As String

    On Error GoTo eh
    
    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If sAspect = vbNullString Then sAspect = Me.Aspect
    If lHKey = 0 Then lHKey = Me.HKey

    If lLocation = 0 Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The location for the session persistent properties is not defined!"
    If lLocation = spp_Registry Then
        If lHKey = 0 Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "The location for the session persistent properties is not defined!"
    End If
    If sSubject = vbNullString Then Err.Raise mErH.AppErr(3), ErrSrc(PROC), "The subject for the session persitent properties is not defined!"
    If sAspect = vbNullString Then Err.Raise mErH.AppErr(4), ErrSrc(PROC), "The aspect under which the session persistent properties are to be stored is not definied!"
    
    If lType = 99 Then lType = spp_REG_SZ
            
    Select Case lLocation
        Case spp_File
            Select Case VarType(vValue)
                Case vbBoolean
                    lChars = WritePrivateProfileString(sAspect, sName, VBA.CStr(VBA.CLng(vValue)), sSubject)
                Case Else
                    lChars = WritePrivateProfileString(sAspect, sName, CStr(vValue), sSubject)
            End Select
            If lChars = 0 Then
                MsgBox "System error when writing property" & vbLf & _
                       "Value  = '" & CStr(vValue) & "'" & vbLf & _
                       "Key    = '" & sName & "'" & vbLf & _
                       "Aspect= '" & sAspect & "'"
            End If
        
        Case spp_Registry
            If Len(vValue) = 0 Then
                NameRemove sName, sAspect
                Exit Sub
            End If
            sPath = HKeyToStrng(lHKey) & "\" & sSubject & "\" & sAspect & "\" & sName
            sRegType = RegTypeStrng(lType)
            Set wss = CreateObject("WScript.Shell")
            wss.RegWrite sPath, vValue, sRegType
            Set wss = Nothing
            
    End Select
    
xt: Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Function Values(Optional lLocation As enLocation = 0, _
                       Optional sSubject As String = vbNullString, _
                       Optional sAspect As String = vbNullString, _
                       Optional lHKey As enHKey = 0) As Dictionary
' -------------------------------------------------------------------
' Returns a Dictionary of values  under a given Aspect (sAspect)
' with the name as the key and the value as the item. When no Aspect
' (sAspect) is provided, the Dictionary returns all names of all
' Aspects of a given Subject.
' -------------------------------------------------------------------
Const PROC      As String = "Values"
Dim dctNames    As Dictionary
Dim v           As Variant
Dim dctValues   As Dictionary
Dim sValue      As String
Dim dctAspects  As Dictionary
Dim vAspect     As Variant
Dim sName       As String

    On Error GoTo eh
    
    Set dctValues = New Dictionary
    Set Values = dctValues
    
    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If lHKey = 0 Then lHKey = Me.HKey
    
    If lLocation = 0 Then Err.Raise mErH.AppErr(1), ErrSrc(PROC), "The location for the session persistent properties is not defined!"
    If lLocation = spp_Registry Then
        If lHKey = 0 Then Err.Raise mErH.AppErr(2), ErrSrc(PROC), "The location for the session persistent properties is not defined!"
    End If
    If sSubject = vbNullString Then Err.Raise mErH.AppErr(3), ErrSrc(PROC), "The subject for the session persistent properties is not defined!"
    If Not Exists(spp_Subjct, lLocation, sSubject, , , lHKey) _
    Then Err.Raise mErH.AppErr(4), ErrSrc(PROC), "The subject '" & sSubject & "' does not exist!"
        
    If sAspect <> vbNullString Then
        If Not Exists(spp_Aspect, lLocation, sSubject, sAspect, , lHKey) _
        Then Err.Raise mErH.AppErr(5), ErrSrc(PROC), "The Aspect '" & sAspect & "' does not exist!"
        '~~> Get the section's keys as an array
        Set dctNames = Names(lLocation, sSubject, sAspect, lHKey)
        For Each v In dctNames
            '~~> Get value of the name v
            sValue = ValueGet(v, , lLocation, sSubject, sAspect, lHKey)
            mBasic.DictAdd dctValues, v, sValue, dct_ascending
        Next v
    Else
        Set dctAspects = Me.Aspects
        For Each vAspect In dctAspects
            sAspect = CStr(vAspect)
            Set dctNames = Names(lLocation, sSubject, sAspect, lHKey)
            For Each v In dctNames
                '~~> Get value of the name v
                sName = dctNames.Item(v)
                sValue = ValueGet(sName, , lLocation, sSubject, vAspect, lHKey)
                mBasic.DictAdd dctValues, sAspect & "." & sName, sValue, dct_ascending
            Next v
        Next vAspect
    End If
    Set Values = dctValues
    
xt: Exit Function
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Function

Public Sub ValuesDisplay(Optional lLocation As enLocation = 0, _
                         Optional sSubject As String = vbNullString, _
                         Optional sAspect As String = vbNullString, _
                         Optional lHKey As enHKey = 0, _
                         Optional ByVal sTitle As String = "Application Properties")
' --------------------------------------------------------------------
'
' --------------------------------------------------------------------
    Const PROC      As String = "Values"
    
    Dim v           As Variant
    Dim sValues     As String
    Dim dctValues   As Dictionary
    Dim sMsg        As tMsg
    On Error GoTo eh
        
    '~~ Use the default values for not provided parameters
    If lLocation = 0 Then lLocation = Me.Location
    If sSubject = vbNullString Then sSubject = Me.Subject
    If lHKey = 0 Then lHKey = Me.HKey

    Set dctValues = Values(lLocation, sSubject, sAspect, lHKey)
    For Each v In dctValues
        sValues = sValues & vbLf & v & " = " & dctValues.Item(v)
    Next v
    sMsg.section(1).sText = sValues
    mMsg.Dsply dsply_title:=sTitle, dsply_message:=sMsg, dsply_buttons:=vbOKOnly
    Exit Sub
    
eh: mErH.ErrMsg ErrSrc(PROC)
End Sub

