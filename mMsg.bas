Attribute VB_Name = "mMsg"
Option Explicit
' -----------------------------------------------------------------------------------------
' Standard Module mMsg Interface for the Common VBA "Alternative" MsgBox (fMsg UserForm)
'
' Methods: Dsply  Exposes all properties and methods for the display of any kind of message
'
' W. Rauschenberger, Berlin Nov 2020
' -----------------------------------------------------------------------------------------
Public Type tMsgSection                 ' ---------------------
       sLabel As String                 ' Structure of the
       sText As String                  ' UserForm's message
       bMonspaced As Boolean            ' area which consists
End Type                                ' of 4 message sections
Public Type tMsg                        ' Attention: 4 is a
       section(1 To 4) As tMsgSection   ' design constant!
End Type                                ' ---------------------

Public Function Box(ByVal dsply_title As String, _
           Optional ByVal dsply_msg As String = vbNullString, _
           Optional ByVal dsply_msg_monospaced As Boolean = False, _
           Optional ByVal dsply_buttons As Variant = vbOKOnly, _
           Optional ByVal dsply_returnindex As Boolean = False, _
           Optional ByVal dsply_min_width As Long = 300, _
           Optional ByVal dsply_max_width As Long = 80, _
           Optional ByVal dsply_max_height As Long = 70, _
           Optional ByVal dsply_min_button_width = 70) As Variant
' -------------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA Message Form as an
' alternative MsgBox.
' Note: In case there is only one single string to be displayed the argument
'       dsply_msg_type will remain unused while the messag is provided via the
'       dsply_msg_strng and dsply_msg_strng_monospaced arguments instead.
'
' See: https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Form.html
'
' W. Rauschenberger, Berlin, Nov 2020
' -------------------------------------------------------------------------------------
    Dim i As Long
    
    With fMsg
        .MaxFormHeightPrcntgOfScreenSize = dsply_max_height ' percentage of screen size
        .MaxFormWidthPrcntgOfScreenSize = dsply_max_width   ' percentage of screen size
        .MinFormWidth = dsply_min_width                     ' defaults to 300 pt. the absolute minimum is 200 pt
        .MinButtonWidth = dsply_min_button_width
        .MsgTitle = dsply_title
        .MsgText(1) = dsply_msg
        .MsgButtons = dsply_buttons
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form improves the performance significantly  ||
        '|| and avoids any flickering message window with its setup.             ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        .Setup '                                                                 ||
        '+------------------------------------------------------------------------+
        .show
    End With
    
    ' -----------------------------------------------------------------------------
    ' Obtaining the reply value/index is only possible when more than one button is
    ' displayed! When the user had a choice the form is hidden when the button is
    ' pressed and the UserForm is unloade when the return value/index (either of
    ' the two) is obtained!
    ' -----------------------------------------------------------------------------
    If dsply_returnindex Then Box = fMsg.ReplyIndex Else Box = fMsg.ReplyValue

End Function

Public Function Buttons(ParamArray dsply_buttons() As Variant) As Collection
' --------------------------------------------------------------------------
' Returns a collection of the items provided by dsply_buttons. When more
' than 7 items are provided the function adds a button row break.
' --------------------------------------------------------------------------
    
    Dim cll As New Collection
    Dim i   As Long
    Dim j   As Long         ' buttons in a row counter
    Dim k   As Long: k = 1  ' button rows counter
    Dim l   As Long         ' total buttons count
    
    On Error Resume Next
    i = LBound(dsply_buttons)
    If Err.Number <> 0 Then GoTo xt
    For i = LBound(dsply_buttons) To UBound(dsply_buttons)
        If (k = 7 And j = 7) Or l = 49 Then GoTo xt
        Select Case dsply_buttons(i)
            Case vbLf, vbCrLf, vbCr
                cll.Add dsply_buttons(i)
                j = 0
                k = k + 1
            Case vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel
                If j = 7 Then
                    cll.Add vbLf
                    j = 0
                    k = k + 1
                End If
                cll.Add dsply_buttons(i)
                j = j + 1
                l = l + 1
            Case Else
                If TypeName(dsply_buttons(i)) = "String" Then
                    ' Any invalid buttons value will be ignored without notice
                    If j = 7 Then
                        cll.Add vbLf
                        j = 0
                        k = k + 1
                    End If
                    cll.Add dsply_buttons(i)
                    j = j + 1
                    l = l + 1
                End If
        End Select
    Next i
    
xt: Set Buttons = cll

End Function
                                     
Public Function Dsply(ByVal dsply_title As String, _
                      ByRef dsply_msg As tMsg, _
             Optional ByVal dsply_buttons As Variant = vbOKOnly, _
             Optional ByVal dsply_returnindex As Boolean = False, _
             Optional ByVal dsply_min_width As Long = 300, _
             Optional ByVal dsply_max_width As Long = 80, _
             Optional ByVal dsply_max_height As Long = 70, _
             Optional ByVal dsply_min_button_width = 70) As Variant
' -------------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA Message Form as an
' alternative MsgBox.
' Note: In case there is only one single string to be displayed the argument
'       dsply_msg_type will remain unused while the messag is provided via the
'       dsply_msg_strng and dsply_msg_strng_monospaced arguments instead.
'
' See: https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Form.html
'
' W. Rauschenberger, Berlin, Nov 2020
' -------------------------------------------------------------------------------------
    Dim i As Long
    
    With fMsg
        .MaxFormHeightPrcntgOfScreenSize = dsply_max_height ' percentage of screen size
        .MaxFormWidthPrcntgOfScreenSize = dsply_max_width   ' percentage of screen size
        .MinFormWidth = dsply_min_width                     ' defaults to 300 pt. the absolute minimum is 200 pt
        .MinButtonWidth = dsply_min_button_width
        .MsgTitle = dsply_title
        For i = 1 To fMsg.NoOfDesignedMsgSections
            .MsgLabel(i) = dsply_msg.section(i).sLabel
            .MsgText(i) = dsply_msg.section(i).sText
            .MsgMonoSpaced(i) = dsply_msg.section(i).bMonspaced
        Next i
        
        .MsgButtons = dsply_buttons
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form improves the performance significantly  ||
        '|| and avoids any flickering message window with its setup.             ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        .Setup '                                                                 ||
        '+------------------------------------------------------------------------+
        .show
    End With
    
    ' -----------------------------------------------------------------------------
    ' Obtaining the reply value/index is only possible when more than one button is
    ' displayed! When the user had a choice the form is hidden when the button is
    ' pressed and the UserForm is unloade when the return value/index (either of
    ' the two) is obtained!
    ' -----------------------------------------------------------------------------
    If dsply_returnindex Then Dsply = fMsg.ReplyIndex Else Dsply = fMsg.ReplyValue

End Function

Public Function ReplyString( _
          ByVal vReply As Variant) As String
' ------------------------------------------
' Returns the Dsply or Box return value as
' string. An invalid value is ignored.
' ------------------------------------------

    If VarType(vReply) = vbString Then
        ReplyString = vReply
    Else
        Select Case vReply
            Case vbAbort:   ReplyString = "Abort"
            Case vbCancel:  ReplyString = "Cancel"
            Case vbIgnore:  ReplyString = "Ignore"
            Case vbNo:      ReplyString = "No"
            Case vbOK:      ReplyString = "Ok"
            Case vbRetry:   ReplyString = "Retry"
            Case vbYes:     ReplyString = "Yes"
        End Select
    End If
    
End Function

