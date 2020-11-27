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

Public Function Dsply(ByVal dsply_title As String, _
                      ByRef dsply_msg_type As tMsg, _
             Optional ByVal dsply_msg_strng As String = vbNullString, _
             Optional ByVal dsply_msg_strng_monospaced As Boolean = False, _
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
        If dsply_msg_strng <> vbNullString Then
            '~~ The message os provided as a simple string
            .MsgText(1) = dsply_msg_strng
        Else
            For i = 1 To fMsg.NoOfDesignedMsgSections
                .MsgLabel(i) = dsply_msg_type.section(i).sLabel
                .MsgText(i) = dsply_msg_type.section(i).sText
                .MsgMonoSpaced(i) = dsply_msg_type.section(i).bMonspaced
            Next i
        End If
        
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


