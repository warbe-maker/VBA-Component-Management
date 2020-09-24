VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12450
   OleObjectBlob   =   "fMsg.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' --------------------------------------------------------------------------
' UserForm fMsg Provides all means for a message with
'               - up to 3 separated text messages, each either with a
'                 proportional or a fixed font
'               - each of the 3 messages with an optional label
'               - 4 reply buttons either specified with replies known
'                 from the VB MsgBox or any test string.
'
' W. Rauschenberger Berlin March 2020
' --------------------------------------------------------------------------
Const H_MARGIN                  As Single = 15
Const V_MARGIN                  As Single = 10
Const MIN_FORM_WIDTH            As Single = 200
Const MIN_REPLY_WIDTH           As Single = 70
Dim sTitle                      As String
Dim sErrSrc                     As String
Dim vReplies                    As Variant
Dim aReplies                    As Variant
Dim sReplies                    As String   ' The provided vReplies converted to reply button values/strings
Dim lNoOfReplyButtons           As Long
Dim siFormWidth                 As Single
Dim sTitleFontName              As String
Dim sTitleFontSize              As String ' Ignored when sTitleFontName is not provided
Dim siTopNext                   As Single
Dim sMsg1Proportional           As String
Dim sMsg2Proportional           As String
Dim sMsg3Proportional           As String
Dim sMsg1Fixed                  As String
Dim sMsg2Fixed                  As String
Dim sMsg3Fixed                  As String
Dim sLabelMessage1              As String
Dim sLabelMessage2              As String
Dim sLabelMessage3              As String
Dim siTitleWidth                As Single
Dim siMaxFixedTextWidth         As Single
Dim siMaxReplyWidth             As Single
Dim siMaxReplyHeight            As Single
Dim vReplyButtons               As Variant
Dim sReplyButtons               As String

Private Sub UserForm_Initialize()
    siFormWidth = MIN_FORM_WIDTH ' Default
End Sub

Public Property Let ErrSrc(ByVal s As String):                  sErrSrc = s:                                    End Property

Public Property Let FormWidth(ByVal si As Single):              siFormWidth = si:                               End Property

Public Property Let LabelMessage1(ByVal s As String):           sLabelMessage1 = s:                             End Property

Public Property Let LabelMessage2(ByVal s As String):           sLabelMessage2 = s:                             End Property

Public Property Let LabelMessage3(ByVal s As String):           sLabelMessage3 = s:                             End Property

Private Property Get LabelMsg1() As MSForms.Label:              Set LabelMsg1 = Me.laMsg1:                      End Property

Private Property Get LabelMsg2() As MSForms.Label:              Set LabelMsg2 = Me.laMsg2:                      End Property

Private Property Get LabelMsg3() As MSForms.Label:              Set LabelMsg3 = Me.laMsg3:                      End Property

Public Property Let Message1Fixed(ByVal s As String):           sMsg1Fixed = s:                                 End Property

Public Property Let Message1Proportional(ByVal s As String):    sMsg1Proportional = s:                          End Property

Public Property Let Message2Fixed(ByVal s As String):           sMsg2Fixed = s:                                 End Property

Public Property Let Message2Proportional(ByVal s As String):    sMsg2Proportional = s:                          End Property

Public Property Let Message3Fixed(ByVal s As String):           sMsg3Fixed = s:                                 End Property

Public Property Let Message3Proportional(ByVal s As String):    sMsg3Proportional = s:                          End Property

Private Property Get Msg1Fixed() As MSForms.TextBox:            Set Msg1Fixed = Me.tbMsg1Fixed:                 End Property

Private Property Get Msg2Fixed() As MSForms.TextBox:            Set Msg2Fixed = Me.tbMsg2Fixed:                 End Property

Private Property Get Msg3Fixed() As MSForms.TextBox:            Set Msg3Fixed = Me.tbMsg3Fixed:                 End Property

Public Property Let Replies(ByVal v As Variant):                vReplies = v:                                   End Property

Public Property Let Title(ByVal s As String):                   sTitle = s:                                     End Property

Public Property Let TitleFontName(ByVal s As String):           sTitleFontName = s:                             End Property

Public Property Let TitleFontSize(ByVal l As Long):             sTitleFontSize = l:                             End Property

Private Property Get TopNext(Optional ByVal ctl As Variant = Nothing) As Single
Dim tb  As MSForms.TextBox
Dim la  As MSForms.Label
Dim cb  As MSForms.CommandButton

    TopNext = siTopNext ' Return the current position for control

    If Not ctl Is Nothing Then
        With ctl            ' Increase the top position for the next control
            .Top = siTopNext
            Select Case TypeName(ctl)
                Case "TextBox"
                    Set tb = ctl
                    siTopNext = tb.Top + tb.Height + V_MARGIN
                Case "CommandButton"
                    Set cb = ctl
                    siTopNext = cb.Top + cb.Height + V_MARGIN
                Case "Label"
                    Set la = ctl
                    Select Case la.Name
                        Case "la"
                            siTopNext = .laTitleSpaceBottom.Top + .laTitleSpaceBottom.Height + V_MARGIN
                        Case Else ' Message label
                            siTopNext = la.Top + la.Height
                    End Select
            End Select
        End With
    End If
End Property

Private Sub AdjustTopPosOfVisibleElements()
    siTopNext = 5   ' initial top position of first visible element
    With Me
        If .laMsg1.Visible Then .laMsg1.Top = TopNext(.laMsg1)
        If .tbMsg1Fixed.Visible Then .tbMsg1Fixed.Top = TopNext(.tbMsg1Fixed)
        If .tbMsg1Proportional.Visible Then .tbMsg1Proportional.Top = TopNext(.tbMsg1Proportional)
        If .laMsg2.Visible Then .laMsg2.Top = TopNext(.laMsg2)
        If .tbMsg2Fixed.Visible Then .tbMsg2Fixed.Top = TopNext(.tbMsg2Fixed)
        If .tbMsg2Proportional.Visible Then .tbMsg2Proportional.Top = TopNext(.tbMsg2Proportional)
        If .laMsg3.Visible Then .laMsg3.Top = TopNext(Me.laMsg3)
        If .tbMsg3Fixed.Visible Then .tbMsg3Fixed.Top = TopNext(.tbMsg3Fixed)
        If .tbMsg3Proportional.Visible Then .tbMsg3Proportional.Top = TopNext(.tbMsg3Proportional)
        
        ReplyButtonsTopPos
        .Height = .cmbReply1.Top + .cmbReply1.Height + (V_MARGIN * 4)
    End With

End Sub

Private Sub cmbReply1_Click():  ReplyWith 0:    End Sub
Private Sub cmbReply2_Click():  ReplyWith 1:    End Sub
Private Sub cmbReply3_Click():  ReplyWith 2:    End Sub
Private Sub cmbReply4_Click():  ReplyWith 3:    End Sub
Private Sub cmbReply5_Click():  ReplyWith 4:    End Sub

Private Sub MsgParagraphFixedWidthReAdjust()
' --------------------------------------------------
' After final adjustment of the forms width
' --------------------------------------------------
Dim siFormWidth As Single

    With Me
        siFormWidth = .Width
        With .tbMsg1Proportional
            If .Visible Then
                .Width = siFormWidth - (V_MARGIN * 2)
            End If
        End With
        With .tbMsg2Proportional
            If .Visible Then
                .Width = siFormWidth - (V_MARGIN * 2)
            End If
        End With
        With .tbMsg3Proportional
            If .Visible Then
                .Width = siFormWidth - (V_MARGIN * 2)
            End If
        End With
        
    End With
End Sub

Private Sub ReplyWith(ByVal lIndex As Long)
' -------------------------------------------------
' Return the value of the clicked reply button.
' -------------------------------------------------
Dim s As String
    
    s = Split(sReplies, ",")(lIndex)
    If IsNumeric(s) Then
        mBasic.MsgReply = CLng(s)
    Else
        mBasic.MsgReply = s
    End If
    Unload Me
End Sub

Private Sub SetupFinalFormWidth()
' ---------------------------------------------
' Final form width adjustment considering:
' - title width
' - maximum fixed message text width
' - width and number of displayed reply buttons
' - specified minimum windo width
' ---------------------------------------------
    Me.Width = Max( _
                   siTitleWidth, _
                   ((siMaxReplyWidth + V_MARGIN) * lNoOfReplyButtons) + (V_MARGIN * 2), _
                   siMaxFixedTextWidth, _
                   MIN_FORM_WIDTH)

End Sub

Private Sub MsgParagraphFixedWidthSet( _
            ByVal tb As MSForms.TextBox, _
            ByVal sText As String)
' ----------------------------------------
' Setup the width of a textbox (tb)
' considering the text (sText) and fixed
' or proportional font (bFixed).
' ----------------------------------------
Dim sSplit      As String
Dim v           As Variant
Dim siMaxWidth  As Single

    '~~ A fixed font Textbox's width is determined by the maximum text line length,
    '~~ determined by means of an autosized width-template
    If InStr(sText, vbLf) <> 0 Then sSplit = vbLf
    If InStr(sText, vbCrLf) <> 0 Then sSplit = vbCrLf
    '~~ Find the width which fits the largest text line
    With Me.tbMsgFixedWidthTemplate
        .MultiLine = False
        .WordWrap = False
        For Each v In Split(sText, sSplit)
            .Value = v
            siMaxWidth = Max(siMaxWidth, .Width)
        Next v
    End With
    
    tb.Width = Max(siMaxWidth, Me.laTitle.Width) + H_MARGIN
    siMaxFixedTextWidth = mBasic.Max(siMaxFixedTextWidth, tb.Width)

End Sub

Private Sub MsgParagraphFixedSetup( _
            ByVal la As MSForms.Label, _
            ByVal sLabelText As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal sTextBoxText As String)
' ----------------------------------------
' Setup any fixed font message and its
' above label when one is specified.
' ----------------------------------------

    If sTextBoxText <> vbNullString Then
        '~~ Setup above text label/title only when there is a text
        If sLabelText <> vbNullString Then
            With la
                .Caption = sLabelText
                .Visible = True
            End With
        End If
        
        With tb
            .Visible = True
            MsgParagraphFixedWidthSet tb, sTextBoxText  ' sets the global siMaxFixedTextWidth variable
            .MultiLine = True
            .WordWrap = True
            .AutoSize = True
            .Value = sTextBoxText
        End With
        
        With Me
            .Width = mBasic.Max(MIN_FORM_WIDTH, _
                                 siFormWidth, _
                                 .laTitle.Width, _
                                 tb.Left + tb.Width + H_MARGIN)
            .laTitle.Width = .Width
            .laTitleSpaceBottom.Width = .Width
        End With
        
    End If
End Sub

Private Sub MsgParagraphProportionalSetup( _
            ByVal la As MSForms.Label, _
            ByVal sLabelText As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal sTextBoxText As String)
' ---------------------------------------
' Adjust message width to form width
' ---------------------------------------
    If sTextBoxText <> vbNullString Then
        '~~ Setup Message Label
        If sLabelText <> vbNullString Then
            With la
                .Caption = sLabelText
                .Visible = True
            End With
        End If
        
        '~~ Setup Message Textbox
        With tb
            .Visible = True
            .MultiLine = True
            .WordWrap = True
            .Width = Me.Width - H_MARGIN
            .AutoSize = True
            .Value = sTextBoxText
        End With
    End If

End Sub

Private Sub MsgParagraphsSetup()
    With Me
        If sMsg1Proportional <> vbNullString _
        Then MsgParagraphProportionalSetup LabelMsg1, sLabelMessage1, .tbMsg2Proportional, sMsg1Proportional
        If sMsg1Fixed <> vbNullString _
        Then MsgParagraphFixedSetup LabelMsg1, sLabelMessage1, Msg1Fixed, sMsg1Fixed
        
        If sMsg2Proportional <> vbNullString _
        Then MsgParagraphProportionalSetup LabelMsg2, sLabelMessage2, .tbMsg2Proportional, sMsg2Proportional
        If sMsg2Fixed <> vbNullString _
        Then MsgParagraphFixedSetup LabelMsg2, sLabelMessage2, Msg2Fixed, sMsg2Fixed
        
        If sMsg3Proportional <> vbNullString _
        Then MsgParagraphProportionalSetup LabelMsg3, sLabelMessage3, .tbMsg3Proportional, sMsg3Proportional
        If sMsg3Fixed <> vbNullString _
        Then MsgParagraphFixedSetup LabelMsg3, sLabelMessage3, Msg3Fixed, sMsg3Fixed
    End With
End Sub

Private Sub ReplyButtonsTopPos()
Dim siTop   As Single

    With Me
        With .cmbReply1
            .Top = TopNext(Me.cmbReply1)
            siTop = .Top
            .Height = siMaxReplyHeight
        End With
        With .cmbReply2
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        With .cmbReply3
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        With .cmbReply4
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        With .cmbReply5
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        .Height = siTop + siMaxReplyHeight + (V_MARGIN * 5)
    End With
    
End Sub

Private Sub ReplyButtonSetup( _
            ByVal cmb As MSForms.CommandButton, _
            ByVal s As String)
' -----------------------------------------------
' Setup Command Button's visibility and text.
' -----------------------------------------------
    If s <> vbNullString Then
        With cmb
            .Visible = True
            .Caption = s
            siMaxReplyHeight = mBasic.Max(siMaxReplyHeight, .Height)
        End With
    End If
End Sub

Private Sub ReplyButtonsSetup(ByVal vReplies As Variant)
' ------------------------------------------------------
' Setup and position the reply buttons. Returns the max
' reply button width.
' ------------------------------------------------------
Dim v   As Variant

    With Me
        '~~ Setup button caption
        If IsNumeric(vReplies) Then
            Select Case vReplies
                Case vbOKOnly
                    lNoOfReplyButtons = 1
                    ReplyButtonSetup .cmbReply1, "Ok"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, MIN_REPLY_WIDTH)
                    sReplies = vbOK
                Case vbOKCancel
                    lNoOfReplyButtons = 2
                    ReplyButtonSetup .cmbReply1, "Ok"
                    ReplyButtonSetup .cmbReply2, "Cancel"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply2.Width, MIN_REPLY_WIDTH)
                    sReplies = vbOK & "," & vbCancel
                Case vbYesNo
                    lNoOfReplyButtons = 2
                    ReplyButtonSetup .cmbReply1, "Yes"
                    ReplyButtonSetup .cmbReply2, "No"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply2.Width, MIN_REPLY_WIDTH)
                    sReplies = vbYes & "," & vbNo
                Case vbRetryCancel
                    lNoOfReplyButtons = 2
                    ReplyButtonSetup .cmbReply1, "Retry"
                    ReplyButtonSetup .cmbReply2, "Cancel"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply2.Width, MIN_REPLY_WIDTH)
                    sReplies = vbRetry & "," & vbCancel
                Case vbYesNoCancel
                    lNoOfReplyButtons = 3
                    ReplyButtonSetup .cmbReply1, "Yes"
                    ReplyButtonSetup .cmbReply2, "No"
                    ReplyButtonSetup .cmbReply3, "Cancel"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply3.Width, MIN_REPLY_WIDTH)
                    sReplies = vbYes & "," & vbNo & "," & vbCancel
                Case vbAbortRetryIgnore
                    lNoOfReplyButtons = 3
                    ReplyButtonSetup .cmbReply1, "Abort"
                    ReplyButtonSetup .cmbReply2, "Retry"
                    ReplyButtonSetup .cmbReply3, "Ignore"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply3.Width, MIN_REPLY_WIDTH)
                    sReplies = vbAbort & "," & vbRetry & "," & vbIgnore
            End Select
        Else
            lNoOfReplyButtons = 0
            sReplies = vbNullString
            aReplies = Split(vReplies, ",")
            For Each v In aReplies
                If v <> vbNullString Then
                    lNoOfReplyButtons = lNoOfReplyButtons + 1
                    sReplies = sReplies & v & ","
                End If
            Next v
            Select Case lNoOfReplyButtons
                Case 1
                    ReplyButtonSetup .cmbReply1, aReplies(0)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, MIN_REPLY_WIDTH)
                Case 2
                    ReplyButtonSetup .cmbReply1, aReplies(0)
                    ReplyButtonSetup .cmbReply2, aReplies(1)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, MIN_REPLY_WIDTH)
                Case 3
                    ReplyButtonSetup .cmbReply1, aReplies(0)
                    ReplyButtonSetup .cmbReply2, aReplies(1)
                    ReplyButtonSetup .cmbReply3, aReplies(2)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, .cmbReply3.Width, MIN_REPLY_WIDTH)
                Case 4
                    ReplyButtonSetup .cmbReply1, aReplies(0)
                    ReplyButtonSetup .cmbReply2, aReplies(1)
                    ReplyButtonSetup .cmbReply3, aReplies(2)
                    ReplyButtonSetup .cmbReply4, aReplies(3)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, .cmbReply3.Width, .cmbReply4.Width, MIN_REPLY_WIDTH)
                Case 5
                    ReplyButtonSetup .cmbReply1, aReplies(0)
                    ReplyButtonSetup .cmbReply2, aReplies(1)
                    ReplyButtonSetup .cmbReply3, aReplies(2)
                    ReplyButtonSetup .cmbReply4, aReplies(3)
                    ReplyButtonSetup .cmbReply5, aReplies(4)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, .cmbReply3.Width, .cmbReply4.Width, .cmbReply5.Width, MIN_REPLY_WIDTH)
            End Select
        End If
    End With

End Sub

Private Sub ReplyButtonsLeftPos()
' --------------------------------------------------------------
' Setup for each reply button its left position.
' --------------------------------------------------------------

    With Me
        Select Case lNoOfReplyButtons
            Case 1
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (siMaxReplyWidth / 2) ' center
                End With
            Case 2
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (V_MARGIN / 2) - siMaxReplyWidth ' left from center
                End With
                With .cmbReply2
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply1.Left + siMaxReplyWidth + V_MARGIN ' right from center
                End With
            Case 3
                With .cmbReply2
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (siMaxReplyWidth / 2) ' center
                End With
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left - siMaxReplyWidth - V_MARGIN ' left from center
                End With
                With .cmbReply3
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left + siMaxReplyWidth + V_MARGIN ' Right from center
                End With
            Case 4
                With .cmbReply2                     ' position 2nd button left from center
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (V_MARGIN / 2) - siMaxReplyWidth
                End With
                With .cmbReply1                 ' position left from button 2
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left - V_MARGIN - siMaxReplyWidth
                End With
                With .cmbReply3
                    .Width = siMaxReplyWidth        ' position 3rd button right from 2nd
                    .Left = Me.cmbReply2.Left + siMaxReplyWidth + V_MARGIN
                End With
                With .cmbReply4
                    .Width = siMaxReplyWidth        ' position 4th button right from 3rd
                    .Left = Me.cmbReply3.Left + siMaxReplyWidth + V_MARGIN
                End With
            Case 5
                With .cmbReply3                                     ' position 3rd reply button in the center
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (siMaxReplyWidth / 2)
                End With
                With .cmbReply2                                     ' position 2nd to left from 3rd
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply3.Left - siMaxReplyWidth - V_MARGIN
                End With
                With .cmbReply1                                     ' position 1st to left from 2nd
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left - siMaxReplyWidth - V_MARGIN
                End With
                With .cmbReply4                                     ' position 4th right from 3rd
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply3.Left + siMaxReplyWidth + V_MARGIN
                End With
                With .cmbReply5                                     ' position 5th right from 4th
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply4.Left + siMaxReplyWidth + V_MARGIN
                End With
        End Select
    End With

End Sub

Private Sub TitleSetup()
' ----------------------------------------------------------------
' When a font name other than the system's font name is provided
' an extra title label mimics the title bar.
' In any case the title label is used to determine the form width
' by autosize of the label.
' ----------------------------------------------------------------
    
    With Me
        If sTitleFontName <> vbNullString And sTitleFontName <> .Font.Name Then
            '~~ A title with a specific font is displayed in a dedicated title label
            With .laTitle   ' Hidden by default
                .Top = TopNext(Me.laTitle)
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .Visible = True
                siTopNext = .Top + .Height + (V_MARGIN / 2)
            End With
            
        Else
            .Caption = sTitle
            .laTitleSpaceBottom.Visible = False
            With .laTitle
                '~~ The title label is used to adjust the form width
                With .Font
                    .Bold = False
                    .Name = Me.Font.Name
                    .Size = 8.6
                End With
                .Visible = False
                siTitleWidth = .Width + H_MARGIN
            End With
            siTopNext = V_MARGIN / 2
        End If
        
        With .laTitle
            '~~ The title label is used to adjust the form width
            With .Font
                .Bold = False
                .Size = 8.6
            End With
            .AutoSize = True
            .Caption = " " & sTitle    ' some left margin
            .AutoSize = False
            siTitleWidth = .Width + H_MARGIN
        End With
        
        .Width = siTitleWidth   ' not the finalwidth though
        .laTitleSpaceBottom.Width = .laTitle.Width
    
    End With

End Sub

Private Sub UserForm_Activate()
    
    With Me
        
        TitleSetup
        
        MsgParagraphsSetup
        
        ReplyButtonsSetup vReplies
        
        SetupFinalFormWidth
                
        MsgParagraphFixedWidthReAdjust
        
        ReplyButtonsLeftPos
        
        AdjustTopPosOfVisibleElements
        
    End With

End Sub

