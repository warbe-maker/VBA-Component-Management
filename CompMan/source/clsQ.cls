VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class Module clsQ: FiFo (queue) services based on a Collection as queue
' ================== serving Null, Empty, Nothing, and all the other type
'                    of variables.
'
' Note: A queue is like a tube which is open at both ends. The first
'       item put into it (add/enqued) is the first one dequeued.
'
' Public services:
' ----------------
' - DeQueue   Returns the first (first added) item from the queue and removes
'             it
' - EnQueue   Queues (adds) an item
' - First     Returns the first item in the queue without dequeuing it.
' - IsEmpty   Returns TRUE when the queue is empty
' - IsInQueue Returns TRUE for a provided and found queue item, the number
'             of items found identical with the provided, and the position of
'             the last one found.
' - Item      Returns an item on a provided position in the queue
' - Last      Returns the last item enqueued
' - Size      Returns the current queue's size
'
' W. Rauschenberger, Berlin Feb 2023
' ----------------------------------------------------------------------------
Private cllQueue As Collection

Private Sub Class_Initialize():                         Set cllQueue = New Collection:          End Sub

Private Sub Class_Terminate():                          Set cllQueue = Nothing:                 End Sub

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

                
Public Sub DeQueue(Optional ByRef d_item_returned As Variant, _
                   Optional ByVal d_item_to_be_dequeued As Variant = Nothing, _
                   Optional ByVal d_item_pos_to_be_dequeued As Long = 0)
' ----------------------------------------------------------------------------
' - When neither a specific item to be dequeued (q_item_to_be_dequeued) nor a
'   specific to be dequeied item by its position (q_item_pos_to_be_dequeued)
'   is provided, the service returns the top item in the queue
'   (q_item_returned) - i.e. the first one added, i.e. enqueued - and removes
'   it from the queue.
' - When a a specific item to be dequeued (q_item_to_be_dequeued) is provided
'   and it exists in the queue, this one is dequeued, returned
'   (q_item_returned) and removed from the queue.
' - When a specific to be dequeued item by its position
'   (q_item_pos_to_be_dequeued) is provided - and the position is within the
'   queue's size - this position's item is returned (d_item_returned) and
'   removed.
' !! When an item to be dequeued (q_item_dequeue) is provided an additionally
' !! specific to be dequeued item by its position (q_item_pos) is ignored.
' ----------------------------------------------------------------------------
        Qdequeue cllQueue, d_item_returned, d_item_to_be_dequeued, d_item_pos_to_be_dequeued
End Sub

Public Sub EnQueue(ByVal e_item As Variant):            Qenqueue cllQueue, e_item:          End Sub

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
    ErrSrc = "clsQ." & sProc
End Function

Public Sub First(ByRef t_item)
    Qfirst cllQueue, t_item
End Sub

Public Function IsEmpty() As Boolean:                   IsEmpty = QisEmpty(cllQueue):       End Function

Public Function IsQueued(ByVal i_item As Variant, _
                Optional ByRef i_pos As Long, _
                Optional ByRef i_no_found As Long) As Boolean
    IsQueued = QisQueued(cllQueue, i_item, i_pos, i_no_found)
End Function

Public Sub Item(ByVal i_pos As Long, _
                ByRef i_item As Variant)
    Qitem cllQueue, i_pos, i_item
End Sub

Public Sub Last(ByRef b_item As Variant):               Qlast cllQueue, b_item:             End Sub

Private Sub Qdequeue(ByRef q_queue As Collection, _
            Optional ByRef q_item_returned As Variant = Nothing, _
            Optional ByVal q_item_to_be_dequeued As Variant = Nothing, _
            Optional ByVal q_item_pos_to_be_dequeued As Long = 0)
' ----------------------------------------------------------------------------
' - When neither a specific item to be dequeued (q_item_to_be_dequeued) nor a
'   specific to be dequed item by its position (q_item_pos) is provided, the
'   service returns the top item in the queue (q_item_returned) - i.e. the
'   first one added, i.e. enqueued - and removes it from the queue.
' - When a specific item to be dequeued (q_item_to_be_dequeued) is provided
'   and it exists in the queue, this one is dequeued, returned
'   (q_item_returned) and removed from the queue.
' - When a specific to be dequeued item by its position (q_item_pos) is
'   provided - and the position is within the queue's size - this position's
'   item is returned and removed.
'
' Notes
' 1. When the argument (q_item_to_be_dequeued) is provided the argument
'    (q_item_pos_to_be_dequeued) is ignored.
' 2. All private procedures Queue... may be copied into any StandardModule
' ----------------------------------------------------------------------------
    Const PROC = "Qdequeue"
    
    On Error GoTo eh
    Dim lPos    As Long
    Dim lNo     As Long
    
    If Not QisEmpty(q_queue) Then
        If Not QisNothing(q_item_to_be_dequeued) Then
            If QisQueued(q_queue, q_item_to_be_dequeued, lPos, lNo) Then
                If lNo > 1 _
                Then Err.Raise AppErr(1), ErrSrc(PROC), "The specific item provided cannot be dequeued since it is not unambigous but in the queue " & lNo & " times!"
                QvarType q_item_to_be_dequeued, q_item_returned
                q_queue.Remove lPos
            End If
        ElseIf q_item_pos_to_be_dequeued <> 0 Then
            If q_item_pos_to_be_dequeued <= Qsize(q_queue) Then
                Qitem q_queue, q_item_pos_to_be_dequeued, q_item_returned
                q_queue.Remove q_item_pos_to_be_dequeued
            End If
        Else
            Qfirst q_queue, q_item_returned
            q_queue.Remove 1
        End If
    End If

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Qenqueue(ByRef q_queue As Collection, _
                     ByVal q_item As Variant)
    If q_queue Is Nothing Then Set q_queue = New Collection
    q_queue.Add q_item
End Sub

Private Sub Qfirst(ByVal q_queue As Collection, _
                   ByRef q_item_returned As Variant)
' ----------------------------------------------------------------------------
' Returns the current first item in the queue, i.e. the one added (enqueued)
' first. When the queue is empty Nothing is returned
' ----------------------------------------------------------------------------
    If Not QisEmpty(q_queue) Then
        QvarType q_queue(1), q_item_returned
    Else
        Set q_item_returned = Nothing
    End If

End Sub

Private Function QidenticalItems(ByVal q_1 As Variant, _
                                 ByVal q_2 As Variant) As Boolean
' ----------------------------------------------------------------------------
' Retunrs TRUE when item 1 is identical with item 2.
' ----------------------------------------------------------------------------
    Select Case True
        Case VarType(q_1) = vbObject And VarType(q_2) = vbObject:   QidenticalItems = q_1 Is q_2
        Case VarType(q_1) <> vbObject And VarType(q_2) <> vbObject: QidenticalItems = q_1 = q_2
    End Select
End Function

Private Function QisEmpty(ByVal q_queue As Collection) As Boolean
    QisEmpty = q_queue Is Nothing
    If Not QisEmpty Then QisEmpty = q_queue.Count = 0
End Function

Private Function QisNothing(ByVal i_item As Variant) As Boolean
    Select Case True
        Case VarType(i_item) = vbNull:      QisNothing = True
        Case VarType(i_item) = vbEmpty:     QisNothing = True
        Case VarType(i_item) = vbObject:    QisNothing = i_item Is Nothing
        Case IsNumeric(i_item):             QisNothing = CInt(i_item) = 0
        Case VarType(i_item) = vbString:    QisNothing = i_item = vbNullString
    End Select
End Function

Private Function QisQueued(ByVal i_queue As Collection, _
                           ByVal i_item As Variant, _
                  Optional ByRef i_pos As Long, _
                  Optional ByRef i_no_found As Long) As Boolean
' ----------------------------------------------------------------------------
' Returns TRUE and the index (i_pos) when the item (i_item) is found in the
' queue (i_queue).
' ----------------------------------------------------------------------------
    Dim i As Long
    
    i_no_found = 0
    For i = 1 To i_queue.Count
        If QidenticalItems(i_queue(i), i_item) Then
            i_no_found = i_no_found + 1
            QisQueued = True
            i_pos = i
        End If
    Next i

End Function

Private Sub Qitem(ByVal q_queue As Collection, _
                  ByVal q_pos As Long, _
         Optional ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (q_item) at the position (q_pos) in the queue (q_queue),
' provided the queue is not empty and the position is within the queue's size.
' ----------------------------------------------------------------------------
    
    If Not QisEmpty(q_queue) Then
        If q_pos <= Qsize(q_queue) Then
            If VarType(q_queue(q_pos)) = vbObject Then
                Set q_item = q_queue(q_pos)
            Else
                q_item = q_queue(q_pos)
            End If
        End If
    End If
    
End Sub

Private Sub Qlast(ByVal q_queue As Collection, _
                  ByRef q_item As Variant)
' ----------------------------------------------------------------------------
' Returns the item (q_item) in the queue which had been enqueued last in the
' queue (q_queue), provided the queue is not empty.
' ----------------------------------------------------------------------------
    Dim lSize As Long
    
    If Not QisEmpty(q_queue) Then
        lSize = q_queue.Count
        If VarType(q_queue(lSize)) = vbObject Then
            Set q_item = q_queue(lSize)
        Else
            q_item = q_queue(lSize)
        End If
    End If
End Sub

Private Function Qsize(ByVal q_queue As Collection) As Long
    If Not QisEmpty(q_queue) Then Qsize = q_queue.Count
End Function

Private Sub QvarType(ByVal q_item As Variant, _
                     ByRef q_item_result As Variant)
' ----------------------------------------------------------------------------
' Returns the pr0vided item (q_item) with respect to its VarType (q_item_var).
' ----------------------------------------------------------------------------
    If VarType(q_item) = vbObject Then
        Set q_item_result = q_item
    Else
        q_item_result = q_item
    End If
End Sub

Public Function Size() As Long:                         Size = Qsize(cllQueue):             End Function

