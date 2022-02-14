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
'               Note: mErH, mTrc, and fMsg are for the mTest module only
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

Private Function AppErr(ByVal lNo As Long) As Long
    AppErr = IIf(lNo < 0, lNo - vbObjectError, vbObjectError + lNo)
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
    Dim bDone           As Boolean
    Dim dctTemp         As Dictionary
    Dim vItem           As Variant
    Dim vItemExisting   As Variant
    Dim vKeyExisting    As Variant
    Dim vValueExisting  As Variant ' the entry's add_key/add_item value for the comparison with the vValueNew
    Dim vValueNew       As Variant ' the argument add_key's/add_item's value
    Dim vValueTarget    As Variant ' the add before/after add_key/add_item's value
    
    On Error GoTo on_error
    
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
            Then Err.Raise mBasic.AppErr(5), ErrSrc(PROC), "The add_target add_key for an add before/after add_key does not exists!"
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
            GoTo end_proc
        End If
        
        '~~ When the add_order is by add_key and not stay with first entry added
        '~~ and the add_key already exists the add_item is updated
        If bOrderByKey And Not add_staywithfirst Then
            If .Exists(add_key) Then
                If VarType(add_item) = vbObject Then Set .Item(add_key) = add_item Else .Item(add_key) = add_item
                GoTo end_proc
            End If
        End If
    End With
        
    '~~ When the add_order argument is an object but does not have a name property raise an error
    If bOrderByKey Then
        If VarType(add_key) = vbObject Then
            On Error Resume Next
            add_key.name = add_key.name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(7), ErrSrc(PROC), "The add_order option is by add_key, the add_key is an object but does not have a name property!"
        End If
    ElseIf bOrderByItem Then
        If VarType(add_item) = vbObject Then
            On Error Resume Next
            add_item.name = add_item.name
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
            GoTo end_proc
        End If
    End With
        
    '~~ Since the new add_key/add_item couldn't simply be added to the Dictionary it will
    '~~ be inserted before or after the add_key/add_item as specified.
    Set dctTemp = New Dictionary
    bDone = False
    
    For Each vKeyExisting In add_dct
        
        If VarType(add_dct.Item(vKeyExisting)) = vbObject _
        Then Set vItemExisting = add_dct.Item(vKeyExisting) _
        Else vItemExisting = add_dct.Item(vKeyExisting)
        
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

end_proc:
    Exit Sub

on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
End Sub

Private Sub ErrMsg(ByVal errnumber As Long, _
                  ByVal errsource As String, _
                  ByVal errdscrptn As String, _
                  ByVal errline As String)
' ----------------------------------------------------
' Display the error message by means of the VBA MsgBox
' ----------------------------------------------------
    
    Dim sErrMsg     As String
    Dim sErrPath    As String
    
    sErrMsg = "Description: " & vbLf & ErrMsgErrDscrptn(errdscrptn) & vbLf & vbLf & _
              "Source:" & vbLf & errsource & ErrMsgErrLine(errline)
    sErrPath = vbNullString ' only available with the mErrHndlr module
    If sErrPath <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Path:" & vbLf & sErrPath
    If ErrMsgInfo(errdscrptn) <> vbNullString _
    Then sErrMsg = sErrMsg & vbLf & vbLf & _
                   "Info:" & vbLf & ErrMsgInfo(errdscrptn)
    MsgBox sErrMsg, vbCritical, ErrMsgErrType(errnumber, errsource) & " in " & errsource & ErrMsgErrLine(errline)

End Sub

Private Function ErrMsgErrType( _
        ByVal errnumber As Long, _
        ByVal errsource As String) As String
' ------------------------------------------
' Return the kind of error considering the
' Err.Source and the error number.
' ------------------------------------------

   If InStr(1, Err.Source, "DAO") <> 0 _
   Or InStr(1, Err.Source, "ODBC Teradata Driver") <> 0 _
   Or InStr(1, Err.Source, "ODBC") <> 0 _
   Or InStr(1, Err.Source, "Oracle") <> 0 Then
      ErrMsgErrType = "Database Error"
   Else
      If errnumber > 0 _
      Then ErrMsgErrType = "VB Runtime Error" _
      Else ErrMsgErrType = "Application Error"
   End If
   
End Function

Private Function ErrMsgErrLine(ByVal errline As Long) As String
    If errline <> 0 _
    Then ErrMsgErrLine = " (at line " & errline & ")" _
    Else ErrMsgErrLine = vbNullString
End Function

Private Function ErrMsgErrDscrptn(ByVal s As String) As String
' -------------------------------------------------------------------
' Return the string before a "||" in the error description. May only
' be the case when the error has been raised by means of err.Raise
' which means when it is an "Application Error".
' -------------------------------------------------------------------
    If InStr(s, DCONCAT) <> 0 _
    Then ErrMsgErrDscrptn = Split(s, DCONCAT)(0) _
    Else ErrMsgErrDscrptn = s
End Function

Private Function ErrMsgInfo(ByVal s As String) As String
' -------------------------------------------------------------------
' Return the string after a "||" in the error description. May only
' be the case when the error has been raised by means of err.Raise
' which means when it is an "Application Error".
' -------------------------------------------------------------------
    If InStr(s, DCONCAT) <> 0 _
    Then ErrMsgInfo = Split(s, DCONCAT)(1) _
    Else ErrMsgInfo = vbNullString
End Function

Private Function DctAddOrderValue(ByVal dctkey As Variant, _
                                  ByVal dctitem As Variant) As Variant
' --------------------------------------------------------------------
' When keyoritem is an object its name becomes the order value else
' the keyoiritem value as is.
' --------------------------------------------------------------------
    If bOrderByKey Then
    
        If VarType(dctkey) = vbObject _
        Then DctAddOrderValue = dctkey.name _
        Else DctAddOrderValue = dctkey
        
    ElseIf bOrderByItem Then
    
        If VarType(dctitem) = vbObject _
        Then DctAddOrderValue = dctitem.name _
        Else DctAddOrderValue = dctitem
    
    End If
    
    If TypeName(DctAddOrderValue) = "String" And bCaseIgnored Then DctAddOrderValue = LCase$(DctAddOrderValue)

End Function

Public Function DctDiffers( _
                ByVal dct1 As Dictionary, _
                ByVal dct2 As Dictionary, _
       Optional ByVal diffitems As Boolean = True, _
       Optional ByVal diffkeys As Boolean = True) As Boolean
' ----------------------------------------------------------
' Returns TRUE when Dictionary 1 (dct1) is different from
' Dictionary 2 (dct2). diffitems and diffkeys may be False
' when only either of the two differences matters.
' When both are false only a differns in the number of
' entries constitutes a difference.
' ----------------------------------------------------------
Const PROC  As String = "DictDiffers"
Dim i       As Long
Dim v       As Variant

    On Error GoTo on_error
    
    '~~ Difference in entries
    DctDiffers = dct1.Count <> dct2.Count
    If DctDiffers Then GoTo exit_proc
    
    If diffitems Then
        '~~ Difference in items
        For i = 0 To dct1.Count - 1
            If VarType(dct1.Items()(i)) = vbObject And VarType(dct1.Items()(i)) = vbObject Then
                DctDiffers = Not dct1.Items()(i) Is dct2.Items()(i)
                If DctDiffers Then GoTo exit_proc
            ElseIf (VarType(dct1.Items()(i)) = vbObject And VarType(dct1.Items()(i)) <> vbObject) _
                Or (VarType(dct1.Items()(i)) <> vbObject And VarType(dct1.Items()(i)) = vbObject) Then
                DctDiffers = True
                GoTo exit_proc
            ElseIf dct1.Items()(i) <> dct2.Items()(i) Then
                DctDiffers = True
                GoTo exit_proc
            End If
        Next i
    End If
    
    If diffkeys Then
        '~~ Difference in keys
        For i = 0 To dct1.Count - 1
            If VarType(dct1.Keys()(i)) = vbObject And VarType(dct1.Keys()(i)) = vbObject Then
                DctDiffers = Not dct1.Keys()(i) Is dct2.Keys()(i)
                If DctDiffers Then GoTo exit_proc
            ElseIf (VarType(dct1.Keys()(i)) = vbObject And VarType(dct1.Keys()(i)) <> vbObject) _
                Or (VarType(dct1.Keys()(i)) <> vbObject And VarType(dct1.Keys()(i)) = vbObject) Then
                DctDiffers = True
                GoTo exit_proc
            ElseIf dct1.Keys()(i) <> dct2.Keys()(i) Then
                DctDiffers = True
                GoTo exit_proc
            End If
        Next i
    End If
       
exit_proc:
    Exit Function
    
on_error:
#If Debugging Then
    Debug.Print Err.Description: Stop: Resume
#End If
    ErrMsg errnumber:=Err.Number, errsource:=ErrSrc(PROC), errdscrptn:=Err.Description, errline:=Erl
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
    ErrSrc = ThisWorkbook.name & " mDct." & sProc
End Function


