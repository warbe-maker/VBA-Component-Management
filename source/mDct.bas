Attribute VB_Name = "mDct"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mDct: Procedures for Dictionaries
'
' Note: 1. Procedures of the mDct module do not use the Common VBA Error Handler.
'          However, test module uses the mErrHndlr module for test purpose.
'
'       2. This module is developed, tested, and maintained in the dedicated
'          Common Component Workbook Dct.xlsm available on Github
'          https://Github.com/warbe-maker/VBA-Basic-Procedures
'
' Methods:
' - DctAdd      Add a key/item pair into a given Dictionary instantly ordered
' - DictDiff    Returns True when tow Dictionaries were different
'
' Uses:         No other modules
'               Note: mErH, mTrc, fMsg, mMsg are for the mTest module only
'
' Requires:     "Microsoft Scripting Runtime"
'               Note: The reference to "Microsoft Visual Basic Application Extensibility .."
'               is for the mTest module only!
'
' W. Rauschenberger, Berlin Sept 2020
' ----------------------------------------------------------------------------
Private bAddedAfter     As Boolean
Private bAddedBefore    As Boolean
Private bCaseIgnored    As Boolean
Private bCaseSensitive  As Boolean
Private bEntrySequence  As Boolean
Private bOrderByItem    As Boolean
Private bOrderByKey     As Boolean
Private bSeqAfterTrgt   As Boolean
Private bSeqAscending   As Boolean
Private bSeqBeforeTrgt  As Boolean
Private bSeqDescending  As Boolean

Public Enum enDctAddOptions ' Dictionary add/insert modes
    sense_caseignored
    sense_casesensitive
    order_byitem
    order_bykey
    seq_aftertarget
    seq_ascending
    seq_beforetarget
    seq_descending
    seq_entry
End Enum

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Sub DctAdd(ByRef add_dct As Dictionary, _
                  ByVal add_key As Variant, _
                  ByVal add_item As Variant, _
         Optional ByVal add_order As enDctAddOptions = order_bykey, _
         Optional ByVal add_seq As enDctAddOptions = seq_entry, _
         Optional ByVal add_sense As enDctAddOptions = sense_casesensitive, _
         Optional ByVal add_target As Variant, _
         Optional ByVal add_staywithfirst As Boolean = False)
' ------------------------------------------------------------------------------------
' Adds the add_item (add_item) to the Dictionary (add_dct) with the add_key (add_key).
' Supports various add_key sequences, case and case insensitive add_key as well
' as adding items before or after an existing add_item.
' - When the add_key (add_key) already exists the add_item is updated when it is
'   numeric or a string, else it is ignored.
' - When the dictionary (add_dct) is Nothing it is setup on the fly.
' - When dctmode = before or after add_target is obligatory and an
'   error is raised when not provided.
' - When the add_item's add_key is an object any dctmode other then by add_seq
'   requires an object with a name property. If not the case an error is
'   raised.

' W. Rauschenberger, Berlin Oct 2020
' ------------------------------------------------------------------------------------
    Const PROC = "DctAdd"
    
    On Error GoTo eh
    Dim bDone           As Boolean
    Dim dctTemp         As Dictionary
    Dim vItem           As Variant
    Dim vItemExisting   As Variant
    Dim vKeyExisting    As Variant
    Dim vValueExisting  As Variant ' the entry's add_key/add_item value for the comparison with the vValueNew
    Dim vValueNew       As Variant ' the argument add_key's/add_item's value
    Dim vValueTarget    As Variant ' the add before/after add_key/add_item's value
    
    If add_dct Is Nothing Then Set add_dct = New Dictionary
    
    '~~ Plausibility checks
    Select Case add_order
        Case order_byitem:  bOrderByItem = True:    bOrderByKey = False
        Case order_bykey:   bOrderByItem = False:   bOrderByKey = True
        Case Else: Err.Raise AppErr(1), ErrSrc(PROC), "Vaue for argument add_order neither add_item nor add_key!"
    End Select
    
    Select Case add_seq
        Case seq_ascending:    bSeqAscending = True:  bSeqDescending = False: bEntrySequence = False: bSeqAfterTrgt = False: bSeqBeforeTrgt = False
        Case seq_descending:   bSeqAscending = False: bSeqDescending = True:  bEntrySequence = False: bSeqAfterTrgt = False: bSeqBeforeTrgt = False
        Case seq_entry:        bSeqAscending = False: bSeqDescending = False: bEntrySequence = True:  bSeqAfterTrgt = False: bSeqBeforeTrgt = False
        Case seq_aftertarget:  bSeqAscending = False: bSeqDescending = False: bEntrySequence = False: bSeqAfterTrgt = True:  bSeqBeforeTrgt = False
        Case seq_beforetarget: bSeqAscending = False: bSeqDescending = False: bEntrySequence = False: bSeqAfterTrgt = False: bSeqBeforeTrgt = True
        Case Else: Err.Raise AppErr(2), ErrSrc(PROC), "Vaue for argument add_seq neither ascending, descendingy, after, before!"
    End Select
    
    Select Case add_sense
        Case sense_caseignored:     bCaseIgnored = True:    bCaseSensitive = False
        Case sense_casesensitive:   bCaseIgnored = False:    bCaseSensitive = True
        Case Else: Err.Raise AppErr(3), ErrSrc(PROC), "Vaue for argument add_sense neither case sensitive nor case ignored!"
    End Select
    
    If bOrderByKey And (bSeqBeforeTrgt Or bSeqAfterTrgt) And add_dct.Exists(add_key) _
    Then Err.Raise AppErr(4), ErrSrc(PROC), "The to be added add_key (add_order value = '" & DctAddOrderValue(add_key, add_item) & "') for an add before/after already exists!"

    '~~ When no add_target is specified for add before/after add_seq descending/ascending is used instead
    If IsMissing(add_target) Then
        If bSeqBeforeTrgt Then add_seq = seq_descending
        If bSeqBeforeTrgt Then add_seq = seq_ascending
    Else
        '~~ When a add_target is specified it must exist
        If (bSeqAfterTrgt Or bSeqBeforeTrgt) And bOrderByKey Then
            If Not add_dct.Exists(add_target) _
            Then Err.Raise AppErr(5), ErrSrc(PROC), "The add_target add_key for an add before/after add_key does not exists!"
        ElseIf (bSeqAfterTrgt Or bSeqBeforeTrgt) And bOrderByItem Then
            If Not DctAddItemExists(add_dct, add_target) _
            Then Err.Raise AppErr(6), ErrSrc(PROC), "The add_target itemfor an add before/after add_item does not exists!"
        End If
        vValueTarget = DctAddOrderValue(add_target, add_target)
    End If
        
    With add_dct
        '~~ When it is the very first add_item or the add_order option
        '~~ is entry sequence the add_item will just be added
        If .Count = 0 Or bEntrySequence Then
            .Add add_key, add_item
            GoTo xt
        End If
        
        '~~ When the order is by key and not stay-with-first-entry-added
        '~~ and the key already exists the item is updated
        If bOrderByKey And Not add_staywithfirst Then
            If .Exists(add_key) Then
                If IsObject(add_item) Then
                    Set .Item(add_key) = add_item
                Else
                    .Item(add_key) = add_item
                End If
                GoTo xt
            End If
        End If
    End With
        
    '~~ When the add_order argument is an object but does not have a name property raise an error
    If bOrderByKey Then
        If IsObject(add_key) Then
            On Error Resume Next
            add_key.Name = add_key.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(7), ErrSrc(PROC), "The add_order option is by add_key, the add_key is an object but does not have a name property!"
        End If
    ElseIf bOrderByItem Then
        If IsObject(add_item) Then
            On Error Resume Next
            add_item.Name = add_item.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(8), ErrSrc(PROC), "The add_order option is by add_item, the add_item is an object but does not have a name property!"
        End If
    End If
    
    vValueNew = DctAddOrderValue(add_key, add_item)
    
    With add_dct
        '~~ Get the last entry's add_order value
        vValueExisting = DctAddOrderValue(.Keys()(.Count - 1), .Items()(.Count - 1))
        
        '~~ When the add_order mode is ascending and the last entry's add_key or add_item
        '~~ is less than the add_order argument just add it and exit
        If bSeqAscending And vValueNew > vValueExisting Then
            .Add add_key, add_item
            GoTo xt
        End If
    End With
        
    '~~ Since the new add_key/add_item couldn't simply be added to the Dictionary it will
    '~~ be inserted before or after the add_key/add_item as specified.
    Set dctTemp = New Dictionary
    bDone = False
    
    For Each vKeyExisting In add_dct
        
        If IsObject(add_dct(vKeyExisting)) _
        Then Set vItemExisting = add_dct(vKeyExisting) _
        Else vItemExisting = add_dct(vKeyExisting)
        
        With dctTemp
            If bDone Then
                '~~ All remaining items just transfer
                .Add vKeyExisting, vItemExisting
            Else
                vValueExisting = DctAddOrderValue(vKeyExisting, vItemExisting)
            
                If vValueExisting = vValueTarget Then
                    If bSeqBeforeTrgt Then
                        '~~ The add before add_target add_key/add_item has been reached
                        .Add add_key, add_item:                     .Add vKeyExisting, vItemExisting:   bDone = True
                        bAddedBefore = True
                    ElseIf bSeqAfterTrgt Then
                        '~~ The add after add_target add_key/add_item has been reached
                        .Add vKeyExisting, vItemExisting:   .Add add_key, add_item:                     bDone = True
                        bAddedAfter = True
                    End If
                ElseIf vValueExisting = vValueNew And bOrderByItem And (bSeqAscending Or bSeqDescending) And Not .Exists(add_key) Then
                    If add_staywithfirst Then
                        .Add vKeyExisting, vItemExisting:   bDone = True ' not added
                    Else
                        '~~ The add_item already exists. When the add_key doesn't exist and add_staywithfirst is False the add_item is added
                        .Add vKeyExisting, vItemExisting:   .Add add_key, add_item:                     bDone = True
                    End If
                ElseIf bSeqAscending And vValueExisting > vValueNew Then
                    .Add add_key, add_item:                     .Add vKeyExisting, vItemExisting:   bDone = True
                ElseIf bSeqDescending And vValueNew > vValueExisting Then
                    '~~> Add before add_target add_key has been reached
                    .Add add_key, add_item:                     .Add vKeyExisting, vItemExisting:   bDone = True
                Else
                    .Add vKeyExisting, vItemExisting ' transfer existing add_item, wait for the one which fits within sequence
                End If
            End If
        End With ' dctTemp
    Next vKeyExisting
    
    '~~ Return the temporary dictionary with the new add_item added and all exiting items in add_dct transfered to it
    '~~ provided ther was not a add before/after error
    If bSeqBeforeTrgt And Not bAddedBefore _
    Then Err.Raise AppErr(9), ErrSrc(PROC), "The add_key/add_item couldn't be added before because the add_target add_key/add_item did not exist!"
    If bSeqAfterTrgt And Not bAddedAfter _
    Then Err.Raise AppErr(10), ErrSrc(PROC), "The add_key/add_item couldn't be added before because the add_target add_key/add_item did not exist!"
    
    Set add_dct = dctTemp
    Set dctTemp = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub AddAscByKey(ByRef add_dct As Dictionary, _
                        ByVal add_key As Variant, _
                        ByVal add_item As Variant)
' ------------------------------------------------------------------------------------
' Adds to the Dictionary (add_dct) an item (add_item) in ascending order by the key
' (add_key). When the key is an object with no Name property an error is raisede.
'
' Note: This is a copy of the DctAdd procedure with fixed options which may be copied
'       into any VBProject's module in order to have it independant from this
'       Common Component.
'
' W. Rauschenberger, Berlin Jan 2022
' ------------------------------------------------------------------------------------
    Const PROC = "AddAscByKey"
    
    On Error GoTo eh
    Dim bDone           As Boolean
    Dim dctTemp         As Dictionary
    Dim vItem           As Variant
    Dim vItemExisting   As Variant
    Dim vKeyExisting    As Variant
    Dim vValueExisting  As Variant ' the entry's add_key/add_item value for the comparison with the vValueNew
    Dim vValueNew       As Variant ' the argument add_key's/add_item's value
    Dim vValueTarget    As Variant ' the add before/after add_key/add_item's value
    Dim bStayWithFirst  As Boolean
    Dim bOrderByItem    As Boolean
    Dim bOrderByKey     As Boolean
    Dim bSeqAscending   As Boolean
    Dim bCaseIgnored    As Boolean
    Dim bCaseSensitive  As Boolean
    Dim bEntrySequence  As Boolean
    
    If add_dct Is Nothing Then Set add_dct = New Dictionary
    
    '~~ Plausibility checks
    bOrderByItem = False
    bOrderByKey = True
    bSeqAscending = True
    bCaseIgnored = False
    bCaseSensitive = True
    bStayWithFirst = True
    bEntrySequence = False
    
    With add_dct
        '~~ When it is the very first add_item or the add_order option
        '~~ is entry sequence the add_item will just be added
        If .Count = 0 Or bEntrySequence Then
            .Add add_key, add_item
            GoTo xt
        End If
        
        '~~ When the add_order is by add_key and not stay with first entry added
        '~~ and the add_key already exists the add_item is updated
        If bOrderByKey And Not bStayWithFirst Then
            If .Exists(add_key) Then
                If IsObject(add_item) Then Set .Item(add_key) = add_item Else .Item(add_key) = add_item
                GoTo xt
            End If
        End If
    End With
        
    '~~ When the add_order argument is an object but does not have a name property raise an error
    If bOrderByKey Then
        If IsObject(add_key) Then
            On Error Resume Next
            add_key.Name = add_key.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(7), ErrSrc(PROC), "The add_order option is by add_key, the add_key is an object but does not have a name property!"
        End If
    ElseIf bOrderByItem Then
        If IsObject(add_item) Then
            On Error Resume Next
            add_item.Name = add_item.Name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(8), ErrSrc(PROC), "The add_order option is by add_item, the add_item is an object but does not have a name property!"
        End If
    End If
    
    vValueNew = DctAddOrderValue(add_key)
    
    With add_dct
        '~~ Get the last entry's add_order value
        vValueExisting = DctAddOrderValue(.Keys()(.Count - 1))
        
        '~~ When the add_order mode is ascending and the last entry's add_key or add_item
        '~~ is less than the add_order argument just add it and exit
        If bSeqAscending And vValueNew > vValueExisting Then
            .Add add_key, add_item
            GoTo xt
        End If
    End With
        
    '~~ Since the new add_key/add_item couldn't simply be added to the Dictionary it will
    '~~ be inserted before or after the add_key/add_item as specified.
    Set dctTemp = New Dictionary
    bDone = False
    
    For Each vKeyExisting In add_dct
        
        If IsObject(add_dct.Item(vKeyExisting)) _
        Then Set vItemExisting = add_dct.Item(vKeyExisting) _
        Else vItemExisting = add_dct.Item(vKeyExisting)
        
        With dctTemp
            If bDone Then
                '~~ All remaining items just transfer
                .Add vKeyExisting, vItemExisting
            Else
                vValueExisting = DctAddOrderValue(vKeyExisting)
            
                If vValueExisting = vValueNew And bOrderByItem And bSeqAscending And Not .Exists(add_key) Then
                    If bStayWithFirst Then
                        .Add vKeyExisting, vItemExisting:   bDone = True ' not added
                    Else
                        '~~ The add_item already exists. When the add_key doesn't exist and bStayWithFirst is False the add_item is added
                        .Add vKeyExisting, vItemExisting:   .Add add_key, add_item:                     bDone = True
                    End If
                ElseIf bSeqAscending And vValueExisting > vValueNew Then
                    .Add add_key, add_item:                     .Add vKeyExisting, vItemExisting:   bDone = True
                Else
                    .Add vKeyExisting, vItemExisting ' transfer existing add_item, wait for the one which fits within sequence
                End If
            End If
        End With ' dctTemp
    Next vKeyExisting
    
    '~~ Return the temporary dictionary with the new add_item added and all exiting items in add_dct transfered to it
    Set add_dct = dctTemp
    Set dctTemp = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub





Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function DctAddOrderValue(ByVal dctkey As Variant, _
                         Optional ByVal dctitem As Variant = Nothing) As Variant
' --------------------------------------------------------------------
' When keyoritem is an object its name becomes the order value else
' the keyoiritem value as is.
' --------------------------------------------------------------------
    If bOrderByKey Then
    
        If IsObject(dctkey) _
        Then DctAddOrderValue = dctkey.Name _
        Else DctAddOrderValue = dctkey
        
    ElseIf bOrderByItem Then
    
        If IsObject(dctitem) _
        Then DctAddOrderValue = dctitem.Name _
        Else DctAddOrderValue = dctitem
    
    End If
    
    If TypeName(DctAddOrderValue) = "String" And bCaseIgnored Then DctAddOrderValue = LCase$(DctAddOrderValue)

End Function

Public Function DctDiffers(ByVal dd_dct1 As Dictionary, _
                           ByVal dd_dct2 As Dictionary, _
                  Optional ByVal dd_diff_items As Boolean = True, _
                  Optional ByVal dd_diff_keys As Boolean = True, _
                  Optional ByVal dd_ignore_items_empty As Boolean = False, _
                  Optional ByVal dd_ignore_case As Boolean = False) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when Dictionary-1 (dd_dct1) differs from Dictionary-2 (dd_dct2).
' - With the option 'Different Items' (dd_diff_items) = TRUE a difference is
'   constituted by different items
' - With the option 'Different Keys' (dd_diff_keys) = TRUE the difference is
'   constituted by different keys
' - When both options are FALSE a difference is constituted by a different number
'   of entries
' Note: When the compared item or key is an object the difference considers by
'       different object names. When the objects do not have a Name property
'       the difference considers different object
' ------------------------------------------------------------------------------
    Const PROC = "DictDiffers"
    
    On Error GoTo eh
    Dim i       As Long
    Dim v       As Variant
    
    If dd_ignore_items_empty Then
        '~~ Remove empty items (items with a lenght = 0) from both Dictionaries
        mDct.RemoveEmptyItems dd_dct1
        mDct.RemoveEmptyItems dd_dct2
    End If
    
    Select Case True
        Case Not dd_diff_items And Not dd_diff_keys
            '~~ A difference is constituted only by a different number of entries
            DctDiffers = dd_dct1.Count <> dd_dct2.Count
        Case dd_diff_items And Not dd_diff_keys
            '~~ A difference is constituted by different items
            For i = 0 To dd_dct1.Count - 1
                If Differs(v1:=dd_dct1.Items()(i) _
                         , v2:=dd_dct2.Items()(i) _
                         , ignore_case:=dd_ignore_case _
                         , ignore_empty:=dd_ignore_items_empty) Then
                    DctDiffers = True
                    Exit For
                End If
            Next i
    
        Case dd_diff_keys And Not dd_diff_items
            '~~ A difference is constituted by different keys
            For i = 0 To dd_dct1.Count - 1
                If Differs(v1:=dd_dct1.Keys()(i) _
                         , v2:=dd_dct2.Keys()(i) _
                         , ignore_case:=dd_ignore_case _
                         , ignore_empty:=dd_ignore_items_empty) Then
                    DctDiffers = True
                    Exit For
                End If
            Next i
    
        Case dd_diff_keys And dd_diff_items
            '~~ A difference is constituted by different keys and different items
            For i = 0 To dd_dct1.Count - 1
                DctDiffers = Differs(v1:=dd_dct1.Keys()(i) _
                                   , v2:=dd_dct2.Keys()(i) _
                                   , ignore_case:=dd_ignore_case _
                                   , ignore_empty:=dd_ignore_items_empty) _
                         And Differs(v1:=dd_dct1.Items()(i) _
                                   , v2:=dd_dct2.Items()(i) _
                                   , ignore_case:=dd_ignore_case _
                                   , ignore_empty:=dd_ignore_items_empty)
                If DctDiffers Then Exit For
            Next i
    End Select
       
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub RemoveEmptyItems(ByRef dct As Dictionary)
' ------------------------------------------------------------------------------
' Removes all empty items (type string lenght = 0) from a Dictionary (dct).
' ------------------------------------------------------------------------------
    Dim v As Variant
    For Each v In dct
        If VarType(dct(v)) = vbString Then
            If Len(Trim(dct(v))) = 0 Then
                dct.Remove v
            End If
        End If
    Next v

End Sub

Private Function Differs(ByVal v1 As Variant, _
                         ByVal v2 As Variant, _
                Optional ByVal ignore_case As Boolean = False, _
                Optional ByVal ignore_empty As Boolean = False) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when v1 is not identical with v2. I.e when they are objects,
' TRUE is returned when the object's Name differ. When only one of the two
' is a string and the other one is an object the string is compared with the
' object's Name property.
' ------------------------------------------------------------------------------
    On Error Resume Next
    Select Case True
        Case IsObject(v1) And IsObject(v2)
            If ignore_case _
            Then Differs = StrComp(v1.Name, v2.Name, vbTextCompare) _
            Else Differs = StrComp(v1.Name, v2.Name, vbBinaryCompare)
            If Err.Number <> 0 Then
                On Error GoTo -1
                Differs = True
            End If
        Case IsObject(v1) And TypeName(v2) = "String"
            If ignore_case _
            Then Differs = StrComp(v1.Name, v2, vbTextCompare) _
            Else Differs = StrComp(v1.Name, v2, vbBinaryCompare)
            If Err.Number <> 0 Then
                On Error GoTo -1
                Differs = True
            End If
        Case TypeName(v1) = "String" And IsObject(v2)
            If ignore_case _
            Then Differs = StrComp(v1, v2.Name, vbTextCompare) _
            Else Differs = StrComp(v1, v2.Name, vbBinaryCompare)
            If Err.Number <> 0 Then
                On Error GoTo -1
                Differs = True
            End If
        Case Else
            If ignore_case _
            Then Differs = StrComp(v1, v2, vbTextCompare) _
            Else Differs = StrComp(v1, v2, vbBinaryCompare)
            If Differs Then
                Debug.Print "Ignore Case = " & CStr(ignore_case)
                Debug.Print "Differs: " & v1 & vbLf & _
                            "         " & v2
            End If
    End Select
End Function

Private Function DctAddItemExists( _
                 ByVal dct As Dictionary, _
                 ByVal dctitem As Variant) As Boolean
    
    Dim v As Variant
    DctAddItemExists = False
    
    For Each v In dct
        If VarType(dct.Item(v)) = vbObject Then
            If dct.Item(v) Is dctitem Then
                DctAddItemExists = True
                Exit Function
            End If
        Else
            If dct.Item(v) = dctitem Then
                DctAddItemExists = True
                Exit Function
            End If
        End If
    Next v
    
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mDct." & sProc
End Function


