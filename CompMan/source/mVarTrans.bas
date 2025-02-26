Attribute VB_Name = "mVarTrans"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mVarTrans: Transforms/translates/transposes variant items
' ========================== such as String, Array, Collection, Dictionary, or
' File) into a String, Array, Collection, Dictionary, or File.
'
' Public procedures:
' ------------------
' ArrayAsCollection      Return an array's items as Collection.
' ArrayAsDictionary      Returns a Dictionary with an array's items as key.
'                        Because the array's items are returned as Directory
'                        keys, the items will become distinct. I e. each item
'                        will exist only once *).
' ArrayAsFile            Writes all items of an array to a file as records
'                        /lines. When no file is provided a temporary file is
'                        written and its name is returned.
' ArrayAsRange           Transferes the content of an array into the range.
' ArrayAsString          Returns a provided array as string with the items
'                        delimited by a vbCrLf.
' CollectionAsArray      Returns a collection's items as array.
' CollectionAsDictionary Returns a collection's items as Dictionary
'                        keys. Because the Collection's items are returned as
'                        Directory keys, the items will become distinct. I. e.
'                        each item will exist only once in the Dictionary *).
' CollectionAsFile       Returns the items of a Collection  as records/lines
'                        in a file, optionally appended.
'                        Uses: StringAsFile, CollectionAsString.
' CollectionAsString     Returns a collection's items as string with the items
'                        delimited by a vbCrLf. Itmes are converted into a
'                        string and if an item is an object its Name property
'                        is used (an error is raised when the object has no
'                        Name property.
' DictionaryAs Array     Returns a Dictionary's keys as array. In case a key
'                        is an object the object's Name property is returned
'                        (an error is raised in case the object has no Name
'                        property).
' DictionaryAsCollection Returns a Collection with a Dictionary's keys as
'                        items.
' DictionaryAsFile       Returns a Dictionary's keys as file records/lines,
'                        optionally appended. Keys are converted into a string,
'                        if an item is an object its Name property is used.
'                        An error is raised when the object has no Name
'                        property.
'                        Uses: StringAsFile, DictionaryAsString.
' DictionaryAsString     Returns a Dictionary's keys as string with each key
'                        delimited by a vbCrLf. Keys are converted into a
'                        string, if an item is an object its Name property is
'                        used. An error is raised when the object has no
'                        Name property.
' FileAsArray            Returns a file's records/lines as array.
' FileAsCollection       Returns a file's records/lines as Collection.
' FileAsDictionary       Returns a file's records/lines as Dictionary keys.
'                        Because the lines become Directory keys, they will
'                        become distinct. I. e. each line will exist only once
'                        in the Dictionary *).
' FileAsString           Returns a file's records/lines as a single string
'                        with the records/lines delimited by a vbCrLf.
'
' ----------------------------------------------------------------------------
' *) To make this restriction productive, the number of occurrences of each
' item is returned as item.
'
' Uses:
' -----
' clsTestAid      Common services supporting test including regression
'                 testing.
'
' Requires:       Reference to "Mocrodoft Scripting Runtime"
'
' W. Rauschenberger, Berlin Jul 2024
' ----------------------------------------------------------------
Private cll          As Collection
Private dct          As Dictionary
Private fso          As New FileSystemObject
Private v            As Variant

Private Property Let Arry(Optional ByRef a_arr As Variant, _
                          Optional ByVal a_indices As Variant = Empty, _
                                   ByVal a_var As Variant)
' ----------------------------------------------------------------------------
' Common WRITE to an array service. The service returns an array (a_arr) with
' the provided item (a_var) either:
' - Simply added, when no index/indices (a_indices) are provided
' - Having added or replaced an item at a given index/indices (a_indices) by
'   concidering that the returned array has new from/to specifics for any of
'   its dimensions at any level whereby the from specific for any dimension
'   remains the same
' - Having created a 1 to 8 dimensions array regarding the provided indices
'   (a_indices) with the provided item added or replaced
' - Having re-dimensioned the provided/input array to the provided indices
'   (a_indices) with the provided item (a_var) added or updated
'
' The index/indices (a_indices) may be provided as:
' - a single integer, indicating that the provided is or the returned will be
'   a 1-dimensional array
' - a string of indices delimited by a comma, indicating that the provided or
'   the returned array is multi-dimensional
' - an array or Collection of indices (for a multi-dimensional array)
'
' Note: In contrast to VBA's ReDim statement this service is able to
'       extend any "to" specification of any dimension (not only the last
'       one) with the only constraint that the "from" specification of any
'       dimension will remain the same.
'
' See ArryReDim for re-specifying any of the dimensions ans even adding new
' dimensions)
'
' Constraints:
' - For a yet not dimensioned and/or not allocated array items may be added
'   by simply not specifying an index
' - For an already dimensioned and/or allocated array the provision of an
'   index for each of its dimensions is obligatory.
'
' Uses: ArryBounds
'
' W. Rauschenberger Berlin, Jan 2025
' ----------------------------------------------------------------------------
    Const PROC = "Arry(Let)"
        
    On Error GoTo eh
    Dim cllSpecNdcs     As New Collection   ' the specified indices as Collection
    Dim cllDimSpecs     As Collection   ' the provided/input array's dimension from/to specifics
    Dim lBase           As Long
    Dim cllBounds       As Collection
    Dim cllBoundsOut    As Collection
    Dim lDimsArry       As Long         ' the provided/input array's number of dimensions
    Dim lDimsSpec       As Long         ' the specifies dimensions derived from the provided indices
    Dim lDimsOut        As Long
    
    lBase = LBound(Array(1))
    
    '~~ Get the provided array's number of dimesion (lDimsArry) and their from/to specifics (cllDimSpecs)
    ArryDims a_arr, cllDimSpecs, lDimsArry      ' DimsArray will be 0 when yet not allocated or not an array
    
    '~~ Get the provided indices as Collection
    Set cllSpecNdcs = ArryIndices(a_indices)
    lDimsSpec = cllSpecNdcs.Count               ' may be 0 when none had been provided
    
    If lDimsArry <> 0 Then
        '~~ The array has at least one Item
        Select Case True
            Case lDimsArry > 1 And lDimsSpec <> lDimsArry
                Err.Raise AppErr(4), ErrSrc(PROC), "For an allocated multidimensional array the provided a_indices are incomplete!"
            
            Case lDimsSpec = 0 And lDimsArry > 1
                '~~ When for an allocated multi-dim array no index has been provided an error is raised
                Err.Raise AppErr(3), ErrSrc(PROC), "To write to a multi-dimensional array an appropriate index is required!"
            
            Case lDimsSpec = 0 And lDimsArry = 1
                '~~ When for an allocated 1-dim array no index is provided, the Item is added to a 1-dim array
                ReDim Preserve a_arr(LBound(a_arr) To UBound(a_arr) + 1)
                a_arr(UBound(a_arr)) = a_var
                
            Case lDimsArry = 1 And cllSpecNdcs(1) > UBound(a_arr)
                '~~ When for an Item in a 1-dim array an index beyond the current specs is provided the array is redimed/epanded accordingly
                ReDim Preserve a_arr(cllDimSpecs(1)(1) To cllSpecNdcs(1))
                a_arr(cllSpecNdcs(1)) = a_var
                
            Case lDimsArry = 1 And cllSpecNdcs(1) <= UBound(a_arr)
                a_arr(cllSpecNdcs(1)) = a_var
            
            Case lDimsArry > 1 And lDimsArry = lDimsSpec
                '~~ The dimensions specified are identical with those of the provided array
                '~~ The dimensios' index may still differ
                ArryBounds a_arr, cllSpecNdcs, cllBoundsOut, cllBounds, lDimsOut
                If lDimsOut = 0 Or lDimsOut = 1 And IsArray(cllBoundsOut(cllBoundsOut.Count)) Then
                    '~~ Either no bounds are out or the out-bound dimension is the last one
                    '~~ which can be re-dimed by VBA's ReDim.
                    Select Case lDimsArry
                        Case 1: ReDim a_arr(cllBounds(1)(1) To cllBounds(1)(2))
                                a_arr(cllSpecNdcs(1)) = a_var
                        
                        Case 2: ReDim a_arr(cllBounds(1)(1) To cllBounds(1)(2), cllBounds(2)(1) To cllBounds(2)(2))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2)) = a_var
                        
                        Case 3: ReDim a_arr(cllBounds(1)(1) To cllBounds(1)(2), cllBounds(2)(1) To cllBounds(2)(2), cllBounds(3)(1) To cllBounds(3)(2))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3)) = a_var
                        
                        Case 4: ReDim a_arr(cllBounds(1)(1) To cllBounds(1)(2), cllBounds(2)(1) To cllBounds(2)(2), cllBounds(3)(1) To cllBounds(3)(2), cllBounds(4)(1) To cllBounds(4)(2))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4)) = a_var
                        
                        Case 5: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5)) = a_var
                        
                        Case 6: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6)) = a_var
                        
                        Case 7: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7)) = a_var
                        
                        Case 8: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8))
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8)) = a_var
                    End Select
                Else
                    Select Case lDimsArry
                        Case 1: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",")
                                a_arr(cllSpecNdcs(1)) = a_var
                        
                        Case 2: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2)) = a_var
                        
                        Case 3: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3)) = a_var
                        
                        Case 4: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4)) = a_var
                        
                        Case 5: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",") _
                                               , "5:" & Join(cllBounds(5), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5)) = a_var
                        
                        Case 6: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",") _
                                               , "5:" & Join(cllBounds(5), ",") _
                                               , "6:" & Join(cllBounds(6), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6)) = a_var
                        
                        Case 7: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",") _
                                               , "5:" & Join(cllBounds(5), ",") _
                                               , "6:" & Join(cllBounds(6), ",") _
                                               , "7:" & Join(cllBounds(7), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7)) = a_var
                        
                        Case 8: ArryReDim a_arr, "1:" & Join(cllBounds(1), ",") _
                                               , "2:" & Join(cllBounds(2), ",") _
                                               , "3:" & Join(cllBounds(3), ",") _
                                               , "4:" & Join(cllBounds(4), ",") _
                                               , "5:" & Join(cllBounds(5), ",") _
                                               , "6:" & Join(cllBounds(6), ",") _
                                               , "7:" & Join(cllBounds(7), ",") _
                                               , "8:" & Join(cllBounds(8), ",")
                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8)) = a_var
                    End Select
                
                End If
            Case Else
                Err.Raise AppErr(6), ErrSrc(PROC), "This is a case of writing to an allocated array yet not considered/implemented!"
        End Select
        
    Else
        '~~ The provided array may yet not have been specified of is empty
        If lDimsSpec = 0 Then
            '~~ Writing to a yet un-allocated or yet un-specified array with
            '~~ no index provided the Item is the first of a new 1-dim array
            ReDim a_arr(lBase To lBase)
            a_arr(lBase) = a_var
        Else
            '~~ For a yet not allocated array an index for a 1- or multi-dimensional array had been provided
            Select Case lDimsSpec
                Case 1: ReDim a_arr(cllSpecNdcs(1))
                        a_arr(cllSpecNdcs(1)) = a_var
                Case 2: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2)):                                                                                        a_arr(cllSpecNdcs(1), cllSpecNdcs(2)) = a_var
                Case 3: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3)):                                                                          a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3)) = a_var
                Case 4: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4)):                                                            a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4)) = a_var
                Case 5: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5)):                                              a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5)) = a_var
                Case 6: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6)):                                a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6)) = a_var
                Case 7: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7)):                  a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7)) = a_var
                Case 8: ReDim a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8)):    a_arr(cllSpecNdcs(1), cllSpecNdcs(2), cllSpecNdcs(3), cllSpecNdcs(4), cllSpecNdcs(5), cllSpecNdcs(6), cllSpecNdcs(7), cllSpecNdcs(8)) = a_var
            End Select
        End If
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get Arry(Optional ByRef a_arr As Variant, _
                          Optional ByVal a_indices As Variant = Nothing) As Variant
' ----------------------------------------------------------------------------
' Common, universal READ from array service supporting up to 8 dimensions.
' The service returns:
' - The Item of a provided array (a_arr) at a given index (a_indices) which
'   might be a single integer or an array or Collection of integers indicating
'   the index for each dimension of a multi-dimensional array.
' - The default (a_default) when
'   - the provided array (a_arr) has no content
'   - the provided array isn't one
'   - the provided index is (indices are) out of the bound/s of any dimension.
'
' W. Rauschenberger Berlin, Jan 2025
' ----------------------------------------------------------------------------
    Const PROC = "Arry(Get)"
    
    Dim lDims  As Long
    Dim bObject As Boolean
    
    Select Case TypeName(a_arr)
        Case "Byte()":     Arry = 0
        Case "Integer()":  Arry = 0
        Case "Long()":     Arry = 0
        Case "Single()":   Arry = 0
        Case "Double()":   Arry = 0
        Case "Currency()": Arry = 0
        Case "Date()":     Arry = #12:00:00 AM#
        Case "String()":   Arry = vbNullString
        Case "Boolean()":  Arry = False
        Case "Variant()":  Arry = Empty
        Case "Object()":   Set Arry = Nothing: bObject = True
        Case Else:         Arry = Empty
    End Select
    
    lDims = ArryDims(a_arr) ' This will return 0 for anything not an array or not a specified array
    
    If lDims > 0 Then
        Set a_indices = ArryIndices(a_indices) ' transforms any kind of provided inedx/indices into a Collection (1 to n)
    End If
    
    On Error Resume Next
    Select Case lDims
        Case 0:
        Case 1: If bObject Then Set Arry = a_arr(a_indices(1)) Else Arry = a_arr(a_indices(1))
        Case 2: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2)) Else Arry = a_arr(a_indices(1), a_indices(2))
        Case 3: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3))
        Case 4: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4))
        Case 5: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5))
        Case 6: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6))
        Case 7: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6), a_indices(7)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6), a_indices(7))
        Case 8: If bObject Then Set Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6), a_indices(7), a_indices(8)) Else Arry = a_arr(a_indices(1), a_indices(2), a_indices(3), a_indices(4), a_indices(5), a_indices(6), a_indices(7), a_indices(8))
    End Select
    
xt:
End Property

Private Function ArryBounds(ByVal a_arr As Variant, _
                            ByVal a_indices As Variant, _
                   Optional ByRef a_out_bounds As Collection, _
                   Optional ByRef a_in_bounds As Collection, _
                   Optional ByRef a_out As Long) As Boolean
' ---------------------------------------------------------------------------
' Returns:
' - TRUE when all dimensions addressed by indices (a_indices) are
'   within the bounds of the respective dimension in array (a_arr)
' - FALSE when any of the provided indices (a_indices) is out of the bounds
'   of the provided array (a_arr)
' - Returns the dimesions which are out of bounds as Collection with items
'   in-bound empty and out-bound with the new bound
' - Returns the complete dimension specifics which combine the "from" spec
'   of the provided array with the new "to" specs in case they are greater
'   than the present ones
'
' Precondition: The indices are provided (a_indices) is either as a single
'               integer - when the array (a_arr) is a 1-dim array - or an
'               array of integers, each specifying the index for one
'               dimension.
'
' Uses: Coll
'
' W. Rauschenberger, Berlin Jan 2025
' ---------------------------------------------------------------------------
    Dim aBounds(1 To 2)     As Variant
    Dim aBoundsOut(1 To 2)  As Variant
    Dim cllBoundsIn         As New Collection
    Dim cllBoundsOut        As New Collection
    Dim cllSpecNdcs         As Collection
    Dim i                   As Long
    Dim lDimsArry           As Long
    Dim lDimsSpec           As Long
    
    lDimsArry = ArryDims(a_arr)
    Set cllSpecNdcs = ArryIndices(a_indices)
    lDimsSpec = cllSpecNdcs.Count
    
    If lDimsSpec > lDimsArry Then GoTo xt
    For i = 1 To cllSpecNdcs.Count
        aBounds(1) = Min(cllSpecNdcs(i), LBound(a_arr, i))
        aBounds(2) = Max(cllSpecNdcs(i), UBound(a_arr, i))
        Coll(cllBoundsIn, i) = aBounds
        If cllSpecNdcs(i) < LBound(a_arr, i) Or cllSpecNdcs(i) > UBound(a_arr, i) Then
            aBoundsOut(1) = LBound(a_arr, i)
            aBoundsOut(2) = UBound(a_arr, i)
            Coll(cllBoundsOut, i) = aBoundsOut
            a_out = a_out + 1
        Else
            Coll(cllBoundsOut, i) = Empty
        End If
    Next i
    
    Set a_in_bounds = cllBoundsIn
    Set a_out_bounds = cllBoundsOut
    ArryBounds = cllBoundsOut.Count > 0
    Set cllBoundsIn = Nothing
    Set cllBoundsOut = Nothing
xt:
End Function

Private Function AppErr(ByVal app_err_no As Long) As Long
' ----------------------------------------------------------------------------
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error. Returns a given
' positive 'Application Error' number (app_err_no) as a negative by adding the
' system constant vbObjectError. Returns the original 'Application Error'
' number when called with a negative error number.
' Obligatory copy Private for any VB-Component using the service but not
' having the mBasic common component installed.
' ----------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Function ArrayAsCollection(ByVal a_arry As Variant) As Collection
' ----------------------------------------------------------------------------
' Return an array's (a_arry) items as Collection.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    With cll
        For Each v In a_arry
            .Add v
        Next v
    End With
    Set ArrayAsCollection = cll
    Set cll = Nothing
    
End Function

Public Function ArrayAsDictionary(ByVal a_arry As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Because the array's items are returned as Directory keys, the
'            items will be unified. I e. each item will exist only once. To
'            make this restriction productive, the number of occurrences of
'            each item is returned as item.'
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    
    With dct
        For Each v In a_arry
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .Item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set ArrayAsDictionary = dct
    Set dct = Nothing
    
End Function

Public Function ArrayAsFile(ByVal a_arry As Variant, _
                   Optional ByRef a_file As Variant = vbNullString, _
                   Optional ByVal a_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes all items of an array (a_arry) to a file (a_file) which might be a
' file object, a file's full name. When no file (a_file) is provided a
' temporary file is returned, else the provided file (a_file) as object.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim sFile As String
    
    On Error GoTo xt ' when no allocated array is provided
    If UBound(a_arry) >= LBound(a_arry) Then
        Select Case True
            Case a_file = vbNullString:         sFile = TempFile
            Case TypeName(a_file) = "File":     sFile = a_file.Path
            Case TypeName(a_file) = "String":   sFile = a_file.Path
        End Select
        
        If a_file_append _
        Then Open sFile For Append As #1 _
        Else Open sFile For Output As #1
        Print #1, Join(a_arry, vbCrLf)
        Close #1
        Set ArrayAsFile = fso.GetFile(sFile)
    End If
    
xt:
End Function

Public Sub ArrayAsRange(ByVal a_arr As Variant, _
                        ByVal a_rng As Range, _
               Optional ByVal a_one_col As Boolean = False)
' ----------------------------------------------------------------------------
' Transferes the content of an array (vArr) into the range (a_rng).
' ----------------------------------------------------------------------------
    Const PROC = "ArryAsRange"
    
    On Error GoTo eh
    Dim rTarget As Range

    If a_one_col Then
        '~~ One column, n rows
        Set rTarget = a_rng.Cells(1, 1).Resize(UBound(a_arr) + 1, 1)
        rTarget.Value = Application.Transpose(a_arr)
    Else
        '~~ One column, n rows
        Set rTarget = a_rng.Cells(1, 1).Resize(1, UBound(a_arr) + 1)
        rTarget.Value = a_arr
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Function ArrayAsString(ByVal a_arry As Variant, _
                     Optional ByVal a_delim As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns an array (a_arry) as string with the items delimited (a_delim).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    ArrayAsString = Join(a_arry, a_delim)
End Function

Private Function ArrayIsAllocated(ByVal a_arry As Variant) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE when an array (a_arry) is allocated.
' ----------------------------------------------------------------------------
    
    On Error Resume Next
    ArrayIsAllocated = UBound(a_arry) >= LBound(a_arry)
    Err.Clear
    On Error GoTo 0
    
End Function

Public Function BooleanAsString(ByVal b As Boolean) As String
    If b Then BooleanAsString = "TRUE" Else BooleanAsString = "FALSE"
End Function

Public Function CollectionAsArray(ByVal c_coll As Collection) As Variant
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr     As Variant
    Dim sName   As String
    
    For Each v In c_coll
        If IsObject(v, sName) _
        Then Arry(arr) = sName _
        Else Arry(arr) = v
    Next v
    CollectionAsArray = arr
    
End Function

Public Function CollectionAsDictionary(ByVal c_coll As Collection) As Dictionary
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as Dictionary keys.
' Attention: Because the collection's items are returned as Directory keys,
'            the items will be unified. I e. each item will exist only once.
'            To make this restriction productive, the number of occurrences of
'            each item is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l As Long
    Dim s As String
    
    With dct
        For Each v In c_coll
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .Item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set CollectionAsDictionary = KeySort(dct)
    Set dct = Nothing

End Function

Public Function CollectionAsFile(ByVal c_coll As Collection, _
                        Optional ByRef c_file_name As String = vbNullString, _
                        Optional ByVal c_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns the items of a Collection (c_coll) as records/lines in a file
' (c_file_name), optionally appended (c_file_append).
' Uses: StringAsFile, CollectionAsString.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    If c_file_name = vbNullString Then c_file_name = TempFile
    StringAsFile CollectionAsString(c_coll), c_file_name, c_file_append
    Set CollectionAsFile = fso.GetFile(c_file_name)

End Function

Public Function CollectionAsString(ByVal c_coll As Collection, _
                          Optional ByRef c_split As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a collection's (c_coll) items as string with the items delimited
' by a vbCrLf. Itmes are converted into a string, if an item is an object its
' Name property is used (an error is raised when the object has no Name
' property.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sName   As String
    Dim sSplit  As String
    Dim v       As Variant
    Dim v2      As Variant
    
    If c_split = vbNullString Then c_split = vbCrLf
    For Each v In c_coll
        Select Case True
            Case IsObject(v, sName)
                s = s & sSplit & sName
                sSplit = c_split
            Case TypeName(v) Like "*()"
                For Each v2 In v
                    s = s & sSplit & CStr(v2)
                    sSplit = c_split
                Next v2
            Case Else
                s = s & sSplit & v
                sSplit = c_split
        End Select
    Next v
    CollectionAsString = s

End Function

Public Function DictionaryAsArray(ByVal d_dict As Dictionary) As Variant
' ----------------------------------------------------------------------------
' Returns a Dictionary's keys as array. In case a key is an object the
' object's Name property is returned (an error is raised in case the object
' has no Name property).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim s   As String
    
    For Each v In d_dict
        If IsObject(v, s) _
        Then Arry(arr) = s _
        Else Arry(arr) = v
    Next v
    DictionaryAsArray = arr
    
End Function

Public Function DictionaryAsCollection(ByVal d_dict As Dictionary) As Collection
' ----------------------------------------------------------------------------
' Returns a Collection with a Dictionary's (d_dict) keys as items.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    With cll
        For Each v In d_dict
            .Add v
        Next v
    End With
    Set DictionaryAsCollection = cll
    Set cll = Nothing
    
End Function

Public Function DictionaryAsDictionary(ByVal d_dict As Dictionary) As Dictionary
' ----------------------------------------------------------------------------
' Returns the Dictionary (d_dict) with the keys in ascending order.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    With dct
        For Each v In d_dict
            .Add v, vbNullString
        Next v
    End With
    Set DictionaryAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Public Function DictionaryAsFile(ByVal d_dict As Dictionary, _
                        Optional ByRef d_file_name As String = vbNullString, _
                        Optional ByVal d_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns a Dictionary's (d_dict) keys as file (d_file_name) records/lines,
' optionally appended (d_file_append). Keys are converted into a string, if an
' item is an object its Name property is used. An error is raised when the
' object has no Name property.
' Uses: StringAsFile, DictionaryAsString.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    If d_file_name = vbNullString Then d_file_name = TempFile
    StringAsFile DictionaryAsString(d_dict), d_file_name, d_file_append
    Set DictionaryAsFile = fso.GetFile(d_file_name)

End Function

Public Function DictionaryAsString(ByVal d_items As Dictionary, _
                          Optional ByRef d_split As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns a Dictionary's keys as string with each key delimited by a vbCrLf.
' Keys are converted into a string, if an item is an object its Name property
' is used. An error is raised when the object has no Name property.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim s       As String
    Dim sName   As String
    Dim sSplit  As String
    
    If d_split = vbNullString Then d_split = vbCrLf
    For Each v In d_items
        If IsObject(v, sName) _
        Then s = s & sSplit & sName _
        Else: s = s & sSplit & v
        sSplit = d_split
    Next v
    DictionaryAsString = s

End Function

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button
' - an "About:" section when the err_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' - the error message either by means of the Common VBA Message Service
'   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
'   means of the VBA.MsgBox in case not.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
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
    
    '~~ Consider extra information is provided with the error description
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
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mVarTrans." & sProc
End Function

Public Function FileAsArray(ByVal f_file As Variant, _
                   Optional ByVal f_empty_excluded = False, _
                   Optional ByVal f_trim As Variant = False) As Variant
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as array.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim sSplit  As String
    Dim s       As String
    
    If TypeName(f_file) = "String" Then f_file = fso.GetFile(f_file)
    s = FileAsString(f_file, sSplit, f_empty_excluded)
    FileAsArray = StringAsArray(s, sSplit, f_trim)
    
End Function

Public Function FileAsCollection(ByVal f_file As Variant) As Collection
' ----------------------------------------------------------------------------
' Returns a file's (f_file) records/lines as Collection.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "FileAsCollection"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim sSplit  As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    With cll
        For Each v In Split(FileAsString(f_file, sSplit), sSplit)
            .Add v
        Next v
    End With
    Set FileAsCollection = cll
    Set cll = Nothing

xt: Exit Function
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FileAsDictionary(ByVal f_file As Variant) As Dictionary
' ----------------------------------------------------------------------------
' Returns a file's records/lines as Dictionary keys.
' Attention: Because the lines become Directory keys, they will become
'            distinct. I. e. each line will exist only once in the Dictionary.
'            To make this restriction productive, the number of occurrences
'            of each line is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct     As New Dictionary
    Dim l       As Long
    Dim s       As String
    Dim sFile   As String
    Dim sSplit  As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    sFile = mVarTrans.FileAsString(f_file, sSplit)
    With dct
        For Each v In Split(sFile, sSplit)
            If Not .Exists(v) Then
                .Add v, 1
            Else
                s = v
                l = .Item(v) + 1
                .Remove v
                .Add s, l
            End If
        Next v
    End With
    Set FileAsDictionary = KeySort(dct)
    Set dct = Nothing
    
End Function

Public Function FileAsFile(ByVal f_file_in As File, _
                           ByVal f_file_out As String, _
                  Optional ByVal f_append As Boolean = False, _
                  Optional ByVal f_rename As Boolean = False) As File
' ----------------------------------------------------------------------------
' Returns a file (f_file_in) as a file with another full name.
' ----------------------------------------------------------------------------
    Const PROC = "FileAsFile"
    
    On Error GoTo eh
    Dim sSplit As String
    
    With fso
        Select Case True
            Case f_rename And Not f_append:     If f_file_in.Path = .GetParentFolderName(f_file_out) _
                                                Then f_file_in.Name = .GetFileName(f_file_out) _
                                                Else Err.Raise AppErr(1), ErrSrc(PROC), "File cannot be renamed when the provided file and the new file's name " & _
                                                                                        "do not point to the same location!"
            Case Not f_rename And Not f_append: .CopyFile f_file_in.Path, f_file_out
            Case Not f_rename And f_append:     StringAsFile FileAsString(f_file_in, sSplit), f_file_out, True
            Case Else:                          Err.Raise AppErr(2), ErrSrc(PROC), "Rename  a n d  append is not supported!"
        End Select
    
        Set FileAsFile = .GetFile(f_file_out)
    End With
    
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function FileAsString(ByVal f_file As Variant, _
                    Optional ByRef f_split As String = vbCrLf, _
                    Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns a file's (f_file) - provided as full name or object - records/lines
' as a single string with the records/lines delimited (f_split).
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Const PROC = "FileAsString"
    
    On Error GoTo eh
    Dim s       As String
    
    If TypeName(f_file) = "File" Then f_file = f_file.Path
    '~~ An error is passed on to the caller
    If Not fso.FileExists(f_file) Then Err.Raise AppErr(1), ErrSrc(PROC), _
                                       "The file, provided by a full name, does not exist!"
    
    Open f_file For Input As #1
    s = Input$(LOF(1), 1)
    Close #1
    
    f_split = SplitIndctr(s) ' may be vbCrLf or vbLf (when file is a download)
    
    '~~ Eliminate any trailing split string
    Do While Right(s, Len(f_split)) = f_split
        s = Left(s, Len(s) - Len(f_split))
        If Len(s) <= Len(f_split) Then Exit Do
    Loop
    
    If f_exclude_empty Then
        s = FileAsStringEmptyExcluded(s)
    End If
    FileAsString = s

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileAsStringEmptyExcluded(ByVal f_s As String) As String
' ----------------------------------------------------------------------------
' Returns a string (f_s) with any empty elements excluded. I.e. the string
' returned begins and ends with a non vbNullString character and has no
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    Do While InStr(f_s, vbCrLf & vbCrLf) <> 0
        f_s = Replace(f_s, vbCrLf & vbCrLf, vbCrLf)
    Loop
    FileAsStringEmptyExcluded = f_s
    
End Function

Private Function IsObject(ByVal i_var As Variant, _
                          ByRef i_name As String) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the object's (i_var) name (i_name) when a variant (i_var)
' is an object. When the object does not have a Name property an error is
' raised.
' ----------------------------------------------------------------------------
    Const PROC = "IsObject"
    
    If Not VBA.IsObject(i_var) Then Exit Function
    IsObject = True
    On Error Resume Next
    i_name = i_var.Name
    If Err.Number <> 0 _
    Then Err.Raise AppErr(1), ErrSrc(PROC), _
         "VarTrans tries to use the Name property of an object when it is to be " & _
         "transferred into a string which is the case when String, Array, or File " & _
         "is the target format. However, the current item is an object which does " & _
         "not have a Name property!"
    
End Function

Private Function KeySort(ByRef k_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (k_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim temp    As Variant
    Dim i       As Long
    Dim j       As Long
    Dim sName   As String
    
    If k_dct Is Nothing Then GoTo xt
    If k_dct.Count = 0 Then GoTo xt
    
    With k_dct
        ReDim arr(0 To .Count - 1)
        For i = 0 To .Count - 1
            If IsObject(.Keys(i), sName) _
            Then arr(i) = sName _
            Else arr(i) = .Keys(i)
        Next i
    End With
    
    '~~ Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(j)
                arr(j) = arr(i)
                arr(i) = temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add key:=vKey, Item:=k_dct.Item(vKey)
    Next i
    
xt: Set k_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function RangeAsArray(ByVal r_rng As Range) As Variant
' ------------------------------------------------------------------------------
' Transferes a range into an array.
' ------------------------------------------------------------------------------
    Const PROC = "RangeAsArray"
    
    Dim arr As Variant
    
    Select Case True
        Case r_rng.Cells.Count = 1:     arr = Array(r_rng.Value)                                            ' single cell
        Case r_rng.Columns.Count = 1:   arr = Application.Transpose(r_rng.Value)                            ' single column
        Case r_rng.Rows.Count = 1:      arr = Application.Transpose(Application.Transpose(r_rng.Value))     ' single row
        Case r_rng.Rows.Count = 2:      arr = r_rng.Value                                                   ' two dimensional array
        Case r_rng.Columns.Count = 2:   arr = r_rng.Value                                                   ' two dimensional array
        Case Else
            Err.Raise AppErr(1), ErrSrc(PROC), "Range cannot be transferred/transposed into an aray!"
    End Select
    RangeAsArray = arr
    
End Function

Public Function SplitIndctr(ByVal s_strng As String, _
                   Optional ByRef s_indctr As String = vbNullString) As String
' ----------------------------------------------------------------------------
' Returns the split indicator of a string (s_strng) as string and as argument
' (s_indctr) provided no split indicator (s_indctr) is already provided.
' The dedection of a split indicator is bypassed in case one has already been
' provided.
' ----------------------------------------------------------------------------
    If s_indctr = vbNullString Then
        Select Case True
            Case InStr(s_strng, vbCrLf) <> 0: s_indctr = vbCrLf
            Case InStr(s_strng, vbLf) <> 0:   s_indctr = vbLf      ' e.g. in case of a downloaded file's_strng complete string
            Case InStr(s_strng, "|&|") <> 0:  s_indctr = "|&|"
            Case InStr(s_strng, ", ") <> 0:   s_indctr = ", "
            Case InStr(s_strng, "; ") <> 0:   s_indctr = "; "
            Case InStr(s_strng, ",") <> 0:    s_indctr = ","
            Case InStr(s_strng, ";") <> 0:    s_indctr = ";"
        End Select
    End If
    SplitIndctr = s_indctr

End Function

Public Function StringAsArray(ByVal s_strng As String, _
                     Optional ByVal s_split As String = vbNullString, _
                     Optional ByVal s_trim As Variant = True) As Variant
' ----------------------------------------------------------------------------
' Returns a string (s_strng) split into an array of strings. When no split
' indicator (s_split) is provided it one is found by examination of the
' string (s_strng). When the option (s_trim) is TRUE (the default), "R", or
' "L" the items in the array are returned trimmed accordingly.
' Example 1: arr = StringAsArray("this is a string", " ") is returned as an
'            array with 3 items: "this", "is", "a", "string".
' Example 2: arr = StringAsArray(FileAsString(FileName),sSplit,False) is
'            returned as any array with records/lines of the provided file,
'            whereby the lines are not trimmed, i.e. leading spaces are
'            preserved.
'            Note: The not provided split indicator has the advantage that it
'                  is provided by the SplitIndctr service, which in that case
'                  returns either vbCrLf or vbLf, the latter when the file is
'                  a download.
' Example 3: arr = FileAsArray(<file>) return the same as example 2.
' Note: Split indicators dedected by examination are: vbCrLf, vbLf, "|&|",
'       ", ", "; ", "," or ";". When neither is dedected vbCrLf is returned.
' ----------------------------------------------------------------------------
    Dim arr As Variant
    Dim i   As Long
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_strng)
    arr = Split(s_strng, SplitIndctr(s_strng, s_split))
    If Not s_trim = False Then
        For i = LBound(arr) To UBound(arr)
            Select Case s_trim
                Case True:  arr(i) = VBA.Trim$(arr(i))
                Case "R":   arr(i) = VBA.RTrim$(arr(i))
                Case "L":   arr(i) = VBA.Trim$(arr(i))
            End Select
        Next i
    End If
    StringAsArray = arr

End Function

Public Function StringAsCollection(ByVal s_items As String, _
                          Optional ByVal s_split As String = vbNullString) As Collection
' ----------------------------------------------------------------------------
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    If s_split = vbNullString Then s_split = SplitIndctr(s_items)
    With cll
        For Each v In Split(s_items, s_split)
            .Add v
        Next v
    End With
    Set StringAsCollection = cll
    Set cll = Nothing
    
End Function

Public Function StringAsDictionary(ByVal s_items As String, _
                          Optional ByVal s_split As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Attention: Transforming the strings within a string (s_items) into a
'            Dictionary by saving the strings as key unifies them. As a
'            compensation of this restriction the number of occurences of a
'            string is returned as item.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    Dim l   As Long
    Dim v   As Variant
    
    Set StringAsDictionary = dct
    If s_items <> vbNullString Then
        If s_split = vbNullString Then s_split = SplitIndctr(s_items)
        With dct
            For Each v In Split(s_items, s_split)
                If Not .Exists(v) Then
                    .Add v, 1
                Else
                    l = dct(v) + 1
                    .Remove v
                    .Add v, l
                End If
            Next v
        End With
        Set StringAsDictionary = KeySort(dct)
    End If
    Set dct = Nothing
    
End Function

Public Function StringAsFile(ByVal s_strng As String, _
                    Optional ByRef s_file As Variant = vbNullString, _
                    Optional ByVal s_file_append As Boolean = False) As File
' ----------------------------------------------------------------------------
' Writes a string (s_strng) to a file (s_file) which might be a file object or
' a file's full name. When no file (s_file) is provided, a temporary file is
' returned.
' Note 1: Only when the string has sub-strings delimited by vbCrLf the string
'         is written a records/lines.
' Note 2: When the string has the alternate split indicator "|&|" this one is
'         replaced by vbCrLf.
' Note when copied: Originates in mVarTrans
'                   See https://github.com/warbe-maker/Excel_VBA_VarTrans
' ----------------------------------------------------------------------------
    
    Select Case True
        Case s_file = vbNullString: s_file = TempFile
        Case TypeName(s_file) = "File": s_file = s_file.Path
    End Select
    
    If s_file_append _
    Then Open s_file For Append As #1 _
    Else Open s_file For Output As #1
    Print #1, s_strng
    Close #1
    Set StringAsFile = fso.GetFile(s_file)
    
End Function

Public Function StringAsString(ByVal v_items As String, _
                      Optional ByRef v_split As String = vbCrLf) As String
' ----------------------------------------------------------------------------
' Returns a string (v_item) with any delimiter string replaced by v_split -
' which defaults to vbCrLf.
' ----------------------------------------------------------------------------
    Const PROC = "StringAsString"
    
    On Error GoTo eh
    StringAsString = Replace(v_items, v_split, v_split)
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function TempFile(Optional ByVal f_path As String = vbNullString, _
                          Optional ByVal f_extension As String = ".txt") As String
' ------------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file. When a path
' (f_path) is omitted in the CurDir path, else in at the provided folder.
' ------------------------------------------------------------------------------
    Dim sTemp As String
    
    If VBA.Left$(f_extension, 1) <> "." Then f_extension = "." & f_extension
    sTemp = Replace(fso.GetTempName, ".tmp", f_extension)
    If f_path = vbNullString Then f_path = fso.GetSpecialFolder(2)
    sTemp = VBA.Replace(f_path & "\" & sTemp, "\\", "\")
    TempFile = sTemp
    fso.CreateTextFile sTemp

End Function

Public Function VarAsArray(ByVal v_items As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "VarAsArray"
    
    Select Case True
        Case TypeName(v_items) = "Collection":  VarAsArray = CollectionAsArray(v_items)
        Case TypeName(v_items) = "File":        VarAsArray = FileAsArray(v_items)
        Case TypeName(v_items) = "Dictionary":  VarAsArray = DictionaryAsArray(v_items)
        Case TypeName(v_items) Like "*()":      VarAsArray = v_items
        Case VarType(v_items) = vbArray:        VarAsArray = v_items
        Case VarType(v_items) = vbString:       VarAsArray = StringAsArray(v_items)
        Case Else:                              Err.Raise AppErr(1), ErrSrc(PROC), _
                                                "The provided v_items argument is of a TypeName """ & TypeName(v_items) & _
                                                """ which is not supported for being transformed into an array!"
    End Select
    
End Function

Public Function VarAsCollection(ByVal v_items As Variant) As Collection
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim cll As New Collection
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsCollection = v_items
        Case TypeName(v_items) = "Dictionary":  Set VarAsCollection = DictionaryAsCollection(v_items)
        Case TypeName(v_items) = "File":        Set VarAsCollection = FileAsCollection(v_items)
        Case TypeName(v_items) Like "*()":      Set VarAsCollection = ArrayAsCollection(v_items)
        Case VarType(v_items) = vbArray:        Set VarAsCollection = ArrayAsCollection(v_items)
        Case VarType(v_items) = vbBoolean:      cll.Add BooleanAsString(v_items)
                                                Set VarAsCollection = New Collection
        Case VarType(v_items) = vbString:       Set VarAsCollection = StringAsCollection(v_items)
        Case Else:                              cll.Add v_items
                                                Set VarAsCollection = cll

    End Select
    
End Function

Public Function VarAsDictionary(ByVal v_items As Variant) As Dictionary
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim dct As New Dictionary
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsDictionary = CollectionAsDictionary(v_items)
        Case TypeName(v_items) = "Dictionary":  Set VarAsDictionary = DictionaryAsDictionary(v_items) ' sort key ascending
        Case TypeName(v_items) = "File":        Set VarAsDictionary = FileAsDictionary(v_items)
        Case TypeName(v_items) = "Nothing":     Set VarAsDictionary = New Dictionary
        Case TypeName(v_items) Like "*()":      Set VarAsDictionary = ArrayAsDictionary(v_items)
        Case VarType(v_items) = vbArray:        Set VarAsDictionary = ArrayAsDictionary(v_items)
        Case VarType(v_items) = vbBoolean:      dct.Add BooleanAsString(v_items), vbNullString
        Case VarType(v_items) = vbString:       Set VarAsDictionary = StringAsDictionary(v_items)
        Case Else:                              Set VarAsDictionary = New Dictionary
    End Select
    
End Function

Public Function VarAsFile(ByVal v_items As Variant, _
                  Optional ByVal v_file As String, _
                  Optional ByVal v_append As Boolean = False) As File
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "VarAsFile"
    
    Select Case True
        Case TypeName(v_items) = "Collection":  Set VarAsFile = CollectionAsFile(v_items, v_file, v_append)
        Case TypeName(v_items) = "File"
        Case TypeName(v_items) = "Dictionary":  Set VarAsFile = DictionaryAsFile(v_items, v_file, v_append)
        Case TypeName(v_items) Like "*()":      Set VarAsFile = ArrayAsFile(v_items, v_file, v_append)
        Case VarType(v_items) = vbArray:        Set VarAsFile = ArrayAsFile(v_items, v_file, v_append)
        Case VarType(v_items) = vbString:       Set VarAsFile = StringAsFile(v_items, v_file, v_append)
        Case Else:                              Err.Raise AppErr(1), ErrSrc(PROC), _
                                                "The provided v_items argument is of a TypeName """ & TypeName(v_items) & _
                                                """ which is not supported for being transformed into an array!"
    End Select
    
End Function

Public Function VarAsString(ByVal v_items As Variant, _
                   Optional ByRef v_split As String) As String
' ----------------------------------------------------------------------------
' Returns variant (v_items) as String whereby the elements are delimited by
' a string (v_split).
' ----------------------------------------------------------------------------

    Select Case True
        Case TypeName(v_items) = "Collection":  VarAsString = CollectionAsString(v_items, v_split)
        Case TypeName(v_items) = "File":        VarAsString = FileAsString(v_items, v_split)
        Case TypeName(v_items) = "Dictionary":  VarAsString = DictionaryAsString(v_items, v_split)
        Case TypeName(v_items) Like "*()":      VarAsString = ArrayAsString(v_items, v_split)
        Case VarType(v_items) = vbArray:        VarAsString = ArrayAsString(v_items, v_split)
        Case VarType(v_items) = vbBoolean:      VarAsString = BooleanAsString(v_items)
        Case VarType(v_items) = vbString:       VarAsString = StringAsString(v_items, v_split)
        Case Else:                              VarAsString = CStr(v_items)
    End Select
    
End Function

