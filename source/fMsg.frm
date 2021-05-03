VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   10560
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   12390
   OleObjectBlob   =   "fMsg.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' -------------------------------------------------------------------------------
' UserForm fMsg Provides all means for a message with up to 5 separated text
'               sections, either proportional- or mono-spaced, with an optional
'               label, and up to 7 reply buttons.
'
' Design:       Since the implementation is merely design driven its setup is
'               essential. Design changes must adhere to the concept.
'
' Uses:         Module mMsg to pass on the clicked reply button to the caller.
'               Note: The UserForm cannot be used directly unless the implemen-
'               tation is mimicked.
'
' Requires:     Reference to "Microsoft Scripting Runtime"
'
' See details at:
' https://warbe-maker.github.io/warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html
'
' W. Rauschenberger Berlin, April 2021 (last revision)
' --------------------------------------------------------------------------
Const DEFAULT_BTTN_MIN_WIDTH            As Single = 70              ' Default minimum reply button width
Const DEFAULT_LBL_MONOSPACED_FONT_NAME  As String = "Courier New"   ' Default monospaced font name
Const DEFAULT_LBL_MONOSPACED_FONT_SIZE  As Single = 9               ' Default monospaced font size
Const DEFAULT_LBL_PROPSPACED_FONT_NAME  As String = "Calibri"       ' Default proportional spaced font name
Const DEFAULT_LBL_PROPSPACED_FONT_SIZE  As Single = 9               ' Default proportional spaced font size
Const DEFAULT_MSG_MAX_HEIGHT_POW        As Long = 70                ' Max form height as a Percentage-Of-Screen-Size
Const DEFAULT_MSG_MAX_WIDTH_POW         As Long = 85                ' Max form width as a percentage of the sreen width
Const DEFAULT_MSG_MIN_WIDTH_PTS         As Single = 400             ' Default minimum message form width
Const DEFAULT_TXT_MONOSPACED_FONT_NAME  As String = "Courier New"   ' Default monospaced font name
Const DEFAULT_TXT_MONOSPACED_FONT_SIZE  As Single = 10              ' Default monospaced font size
Const DEFAULT_TXT_PROPSPACED_FONT_NAME  As String = "Tahoma"        ' Default proportional spaced font name
Const DEFAULT_TXT_PROPSPACED_FONT_SIZE  As Single = 10              ' Default proportional spaced font size
Const NEXT_ROW                          As String = vbLf            ' Reply button row break
Const SPACE_HORIZONTAL_BTTN_AREA        As Single = 10              ' Minimum left and right margin for the centered buttons area
Const SPACE_HORIZONTAL_BUTTONS          As Single = 4               ' Horizontal space between reply buttons
Const SPACE_HORIZONTAL_LEFT             As Single = 0               ' Left margin for labels and text boxes
Const SPACE_HORIZONTAL_RIGHT            As Single = 15              ' Horizontal right space for labels and text boxes
Const SPACE_HORIZONTAL_SCROLLBAR        As Single = 18              ' Additional horizontal space required for a frame with a vertical scrollbar
Const SPACE_VERTICAL_AREAS              As Single = 10              ' Vertical space between message area and replies area
Const SPACE_VERTICAL_BOTTOM             As Single = 50              ' Bottom space after the last displayed reply row
Const SPACE_VERTICAL_BTTN_ROWS          As Single = 5               ' Vertical space between button rows
Const SPACE_VERTICAL_LABEL              As Single = 0               ' Vertical space between label and the following text
Const SPACE_VERTICAL_SCROLLBAR          As Single = 12              ' Additional vertical space required for a frame with a horizontal scroll barr
Const SPACE_VERTICAL_SECTIONS           As Single = 5               ' Vertical space between displayed message sections
Const SPACE_VERTICAL_TEXTBOXES          As Single = 18              ' Vertical bottom marging for all textboxes
Const SPACE_VERTICAL_TOP                As Single = 2               ' Top position for the first displayed control
Const TEST_WITH_FRAME_BORDERS           As Boolean = False          ' For test purpose only! Display frames with visible border
Const TEST_WITH_FRAME_CAPTIONS          As Boolean = False          ' For test purpose only! Display frames with their test captions (erased by default)

' ------------------------------------------------------------
' Means to get and calculate the display devices DPI in points
Const SM_XVIRTUALSCREEN                 As Long = &H4C&
Const SM_YVIRTUALSCREEN                 As Long = &H4D&
Const SM_CXVIRTUALSCREEN                As Long = &H4E&
Const SM_CYVIRTUALSCREEN                As Long = &H4F&
Const LOGPIXELSX                        As Long = 88
Const LOGPIXELSY                        As Long = 90
Const TWIPSPERINCH                      As Long = 1440
Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
' ------------------------------------------------------------

Private Enum enStartupPosition      ' ---------------------------
    sup_Manual = 0                  ' Used to position the
'    sup_CenterOwner = 1             ' final setup message form
'    sup_CenterScreen = 2            ' horizontally and vertically
'    sup_WindowsDefault = 3          ' centered on the screen
End Enum                            ' ---------------------------

Dim bDoneHeightDecrement            As Boolean
Dim bDoneMonoSpacedSections         As Boolean
Dim bDoneMsgArea                    As Boolean
Dim bDonePropSpacedSections         As Boolean
Dim bDoneSetup                      As Boolean
Dim bDoneTitle                      As Boolean
Dim bDsplyFrmsWthBrdrsTestOnly      As Boolean
Dim bDsplyFrmsWthCptnTestOnly       As Boolean
Dim bProgressMode                   As Boolean
Dim bFormEvents                     As Boolean
Dim bHscrollbarBttnsArea            As Boolean
Dim bVscrollbarBttnsArea            As Boolean
Dim bVscrollbarMsgArea              As Boolean
Dim cllDsgnAreas                    As Collection   ' Collection of the two primary/top frames
Dim cllDsgnBttnRows                 As Collection   ' Collection of the designed reply button row frames
Dim cllDsgnBttns                    As Collection   ' Collection of the collection of the designed reply buttons of a certain row
Dim cllDsgnBttnsFrame               As Collection
Dim cllDsgnRowBttns                 As Collection   ' Collection of a designed reply button row's buttons
Dim cllDsgnSections                 As Collection   '
Dim cllDsgnSectionsLabel            As Collection
Dim cllDsgnSectionsText             As Collection   ' Collection of section frames
Dim cllDsgnSectionsTextFrame        As Collection
Dim AppliedBttnRows                 As Dictionary   ' Dictionary of applied/used/visible button rows (key=frame, item=row)
Dim AppliedBttns                    As Dictionary   ' Dictionary of applied buttons (key=CommandButton, item=row)
Dim AppliedBttnsRetVal              As Dictionary   ' Dictionary of the applied buttons' reply value (key=CommandButton)
Dim dctAppliedControls              As Dictionary   ' Dictionary of all applied controls (versus just designed)
Dim dctSectionsLabel                As Dictionary   ' Section specific label either provided via properties MsgLabel or Msg
Dim dctSectionsMonoSpaced           As Dictionary   ' Section specific monospace option either provided via properties MsgMonospaced or Msg
Dim dctSectionsText                 As Dictionary   ' Section specific text either provided via properties MsgText or Msg
Dim lMaxMsgHeightPoSh               As Long         ' Maximum % of the screen height
Dim lMaxMsgWidthPoSw                As Long         ' Maximum % of the screen width
Dim lMinMsgWidthPoSw                As Long         ' Minimum % of the screen width - calculated when min form width in pt is assigend
Dim lMinMsgHeightPoSw               As Long
Dim lMinMsgHeightPoSh               As Long         ' Minimum % of the screen width - calculated when min form width in pt is assigend
Dim lReplyIndex                     As Long         ' Index of the clicked reply button (a value ranging from 1 to 49)
Dim lSetupRowButtons                As Long         ' number of buttons setup in a row
Dim lSetupRows                      As Long         ' number of setup button rows
Dim siHmarginButtons                As Single
Dim siHmarginFrames                 As Single       ' Test property, value defaults to 0
Dim siMaxButtonHeight               As Single
Dim siBttnsFrameMaxWidth            As Single
Dim siMaxButtonWidth                As Single
Dim siMaxMsgHeightPts               As Single       ' above converted to excel userform height
Dim siMaxMsgWidthPts                As Single       ' above converted to excel userform width
Dim siMaxSectionWidth               As Single
Dim siMinButtonWidth                As Single
Dim siMinMsgHeightPts               As Single
Dim siMinMsgWidthPts                As Single
Dim siTitleWidth                    As Single
Dim siVmarginButtons                As Single
Dim siVmarginFrames                 As Single       ' Test property, value defaults to 0
Dim sMonoSpacedLabelDefaultFontName As String
Dim sMonoSpacedLabelDefaultFontSize As Single
Dim sMonoSpacedTextDefaultFontName  As String
Dim sMonoSpacedTextDefaultFontSize  As Single
Dim sTitle                          As String
Dim sTitleFontName                  As String
Dim sTitleFontSize                  As String       ' Ignored when sTitleFontName is not provided
Dim vbuttons                        As Variant
Dim vDefaultBttn                    As Variant      ' Index or caption of the default button
Dim vReplyValue                     As Variant
Dim VirtualScreenHeightPts          As Single
Dim VirtualScreenWidthPts           As Single
Dim VirtualScreenLeftPts            As Single
Dim VirtualScreenTopPts             As Single
Dim bReplyWithIndex                 As Boolean

Private Sub UserForm_Initialize()
    Const PROC = "UserForm_Initialize"
    
    On Error GoTo eh
    siMinButtonWidth = DEFAULT_BTTN_MIN_WIDTH
    siHmarginButtons = SPACE_HORIZONTAL_BUTTONS
    siVmarginButtons = SPACE_VERTICAL_BTTN_ROWS
    bFormEvents = False
    GetScreenMetrics VirtualScreenLeftPts, VirtualScreenTopPts, _
                     VirtualScreenWidthPts, VirtualScreenHeightPts      ' This environment screen's width and height in pts
    Me.MinMsgWidthPts = DEFAULT_MSG_MIN_WIDTH_PTS                       ' computes the min width in % of sreen width)
    Me.MaxMsgWidthPrcntgOfScreenSize = DEFAULT_MSG_MAX_WIDTH_POW        ' computes the max width in pts (siMaxMsgWidthPts)
    Me.MaxMsgHeightPrcntgOfScreenSize = DEFAULT_MSG_MAX_HEIGHT_POW      ' computes the max height in pts (siMaxMsgHeightPts)
    sMonoSpacedTextDefaultFontName = DEFAULT_TXT_MONOSPACED_FONT_NAME
    sMonoSpacedTextDefaultFontSize = DEFAULT_TXT_MONOSPACED_FONT_SIZE
    sMonoSpacedLabelDefaultFontName = DEFAULT_LBL_MONOSPACED_FONT_NAME
    sMonoSpacedLabelDefaultFontSize = DEFAULT_LBL_MONOSPACED_FONT_SIZE
    bDsplyFrmsWthCptnTestOnly = False
    bDsplyFrmsWthBrdrsTestOnly = False
    FormWidth = Me.MinMsgWidthPts
    Me.Height = 60          ' let's start with this
    siHmarginFrames = 0     ' Ensures proper command buttons framing, may be used for test purpose
    Me.VmarginFrames = 0    ' Ensures proper command buttons framing and vertical positioning of controls
    bDoneSetup = False
    bDoneTitle = False
    bDoneMonoSpacedSections = False
    bDonePropSpacedSections = False
    bDoneMsgArea = False
    bDoneHeightDecrement = False
    vDefaultBttn = 1
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub UserForm_Terminate()
    Set cllDsgnAreas = Nothing
    Set cllDsgnBttnRows = Nothing
    Set cllDsgnBttns = Nothing
    Set cllDsgnBttnsFrame = Nothing
    Set cllDsgnRowBttns = Nothing
    Set cllDsgnSections = Nothing
    Set cllDsgnSectionsLabel = Nothing
    Set cllDsgnSectionsText = Nothing
    Set cllDsgnSectionsTextFrame = Nothing
    Set AppliedBttnRows = Nothing
    Set AppliedBttns = Nothing
    Set AppliedBttnsRetVal = Nothing
    Set dctAppliedControls = Nothing
    Set dctSectionsLabel = Nothing
    Set dctSectionsMonoSpaced = Nothing
    Set dctSectionsText = Nothing
End Sub

Public Property Let ProgressMode(ByVal b As Boolean)
    bProgressMode = b
End Property

Private Property Get AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton) As Variant
    AppliedButtonRetVal = AppliedBttnsRetVal(Button)
End Property

Private Property Let AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton, ByVal v As Variant)
    AppliedBttnsRetVal.Add Button, v
End Property

Private Property Get AppliedButtonRowHeight() As Single
    AppliedButtonRowHeight = siMaxButtonHeight + (siVmarginFrames * 2) + 2
End Property

Private Property Let AppliedControls(ByVal v As Variant)
    If dctAppliedControls Is Nothing Then Set dctAppliedControls = New Dictionary
    If Not IsApplied(v) Then dctAppliedControls.Add v, v.name
End Property

Private Property Get BttnsAreaWidthUsable() As Single
    BttnsAreaWidthUsable = Me.InsideWidth - SPACE_HORIZONTAL_BTTN_AREA
End Property

Private Property Get BttnsFrameHeight() As Single
    Dim l As Long:  l = AppliedBttnRows.Count
    BttnsFrameHeight = (AppliedButtonRowHeight * l) + (siVmarginButtons * (l - 1)) + (siVmarginFrames * 2) + 2
End Property

Private Property Get ClickedButtonIndex(Optional ByVal cmb As MSForms.CommandButton) As Long
    
    Dim i   As Long
    Dim v   As Variant
    
    For Each v In AppliedBttnsRetVal
        i = i + 1
        If v Is cmb Then
            ClickedButtonIndex = i
            Exit For
        End If
    Next v

End Property

Public Property Get ReplyValue() As Variant:    ReplyValue = vReplyValue:   End Property
Public Property Let DefaultBttn(ByVal vDefault As Variant)
    vDefaultBttn = vDefault
End Property

Private Property Get DsgnBttn(Optional ByVal bttn_row As Long, Optional ByVal bttn_no As Long) As MSForms.CommandButton
    Set DsgnBttn = cllDsgnBttns(bttn_row)(bttn_no)
End Property

Private Property Get DsgnBttnRow(Optional ByVal row As Long) As MSForms.Frame:            Set DsgnBttnRow = cllDsgnBttnRows(row):                         End Property

Private Property Get DsgnBttnRows() As Collection:                                        Set DsgnBttnRows = cllDsgnBttnRows:                             End Property

Private Property Get DsgnBttnsArea() As MSForms.Frame:                                    Set DsgnBttnsArea = cllDsgnAreas(2):                              End Property

Private Property Get DsgnBttnsFrame() As MSForms.Frame:                                   Set DsgnBttnsFrame = cllDsgnBttnsFrame(1):                      End Property

Private Property Get DsgnMsgArea() As MSForms.Frame:                                        Set DsgnMsgArea = cllDsgnAreas(1):                                  End Property

Private Property Get DsgnSection(Optional msg_section As Long) As MSForms.Frame:            Set DsgnSection = cllDsgnSections(msg_section):                     End Property

Private Property Get DsgnSectionLabel(Optional msg_section As Long) As MSForms.Label:       Set DsgnSectionLabel = cllDsgnSectionsLabel(msg_section):           End Property

Private Property Get DsgnSections() As Collection:                                          Set DsgnSections = cllDsgnSections:                                 End Property

Private Property Get DsgnSectionText(Optional msg_section As Long) As MSForms.TextBox:      Set DsgnSectionText = cllDsgnSectionsText(msg_section):             End Property

Private Property Get DsgnSectionTextFrame(Optional ByVal msg_section As Long):              Set DsgnSectionTextFrame = cllDsgnSectionsTextFrame(msg_section):   End Property

Private Property Get DsgnTextFrame(Optional ByVal msg_section As Long) As MSForms.Frame:    Set DsgnTextFrame = cllDsgnSectionsTextFrame(msg_section):          End Property

Private Property Get DsgnTextFrames() As Collection:                                        Set DsgnTextFrames = cllDsgnSectionsTextFrame:                      End Property

Public Property Let DsplyFrmsWthBrdrsTestOnly(ByVal b As Boolean)
    
    Dim ctl As MSForms.Control
       
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Or TypeName(ctl) = "TextBox" Then
            ctl.BorderColor = -2147483638   ' active frame, allows with style none to hide the frame
            If b = False _
            Then ctl.BorderStyle = fmBorderStyleNone _
            Else ctl.BorderStyle = fmBorderStyleSingle
        End If
    Next ctl
    
End Property

Public Property Let DsplyFrmsWthCptnTestOnly(ByVal b As Boolean):                           bDsplyFrmsWthCptnTestOnly = b:                                      End Property

Private Property Let FormWidth(ByVal w As Single)
    Dim siInOutDiff As Single:  siInOutDiff = Me.Width - Me.InsideWidth
    Me.Width = Max(Me.Width, siTitleWidth, siMinMsgWidthPts, w + siInOutDiff)
    Me.Width = Min(Me.Width, siMaxMsgWidthPts)
End Property

Private Property Let HeightDecrementBttnsArea(ByVal b As Boolean)
    bVscrollbarBttnsArea = b
    bDoneHeightDecrement = b
End Property

Private Property Let HeightDecrementMsgArea(ByVal b As Boolean)
    bVscrollbarMsgArea = b
    bDoneHeightDecrement = b
End Property

Public Property Let HmarginButtons(ByVal si As Single):                                 siHmarginButtons = si:                                                      End Property

Public Property Let HmarginFrames(ByVal si As Single):                                  siHmarginFrames = si:                                                       End Property

Private Property Get IsApplied(Optional ByVal v As Variant) As Boolean
    If dctAppliedControls Is Nothing _
    Then IsApplied = False _
    Else IsApplied = dctAppliedControls.Exists(v)
End Property

Public Property Get MaxMsgHeightPrcntgOfScreenSize() As Long:                           MaxMsgHeightPrcntgOfScreenSize = lMaxMsgHeightPoSh:                        End Property

Public Property Let MaxMsgHeightPrcntgOfScreenSize(ByVal l As Long)
' ------------------------------------------------------------------
' The maximum message height must not exceed 70 % of the screen size
' ------------------------------------------------------------------
    lMaxMsgHeightPoSh = l
    siMaxMsgHeightPts = VirtualScreenHeightPts * (Min(l, 70) / 100)
End Property

Public Property Get MaxMsgHeightPts() As Single
    MaxMsgHeightPts = siMaxMsgHeightPts
End Property

Public Property Get MaxMsgWidthPrcntgOfScreenSize() As Long:                            MaxMsgWidthPrcntgOfScreenSize = lMaxMsgWidthPoSw:                          End Property

Public Property Let MaxMsgWidthPrcntgOfScreenSize(ByVal l As Long)
' ------------------------------------------------------------------------------
' Determines the maximum width of the Message-Form as a percentage of the screen
' width considering that the percentage cannot exceed 99% of the screen widht
' and it cannot be less than the specified minimum Message-Form width.
' ------------------------------------------------------------------------------
    lMaxMsgWidthPoSw = Max(lMinMsgHeightPoSh, l)
    siMaxMsgWidthPts = VirtualScreenWidthPts * (Min(lMaxMsgWidthPoSw, 99) / 100)
End Property

Public Property Get MaxMsgWidthPts() As Single
' --------------------------------------------
' For display with test procedures
' --------------------------------------------
    MaxMsgWidthPts = siMaxMsgWidthPts
End Property

Private Property Get MaxRowsHeight() As Single:                                         MaxRowsHeight = siMaxButtonHeight + (siVmarginFrames * 2):                  End Property

Public Property Let MinButtonWidth(ByVal si As Single):                                 siMinButtonWidth = si:                                                      End Property

Public Property Get MinMsgWidthPoSw() As Long:                                          MinMsgWidthPoSw = lMinMsgWidthPoSw:                                  End Property
Public Property Get MinMsgHeightPoSw() As Long:                                         MinMsgHeightPoSw = lMinMsgHeightPoSw:                                  End Property

Public Property Get MinMsgWidthPts() As Single:                                         MinMsgWidthPts = siMinMsgWidthPts:                                          End Property
Public Property Get MinMsgHeightPts() As Single:                                        MinMsgWidthPts = siMinMsgHeightPts:                                          End Property

Public Property Let MinMsgWidthPts(ByVal si As Single)
' ----------------------------------------------------------------
' The maximum message width must not become less than the minimum.
' ----------------------------------------------------------------
    siMinMsgWidthPts = Max(si, 200) ' cannot be specified less
    If siMaxMsgWidthPts < siMinMsgWidthPts Then
       siMaxMsgWidthPts = siMinMsgWidthPts
    End If
    lMinMsgHeightPoSh = CInt((siMinMsgWidthPts / VirtualScreenWidthPts) * 100)
End Property

Public Property Let MsgButtons(ByVal v As Variant)
        
    Select Case VarType(v)
        Case VarType(v) = vbLong, vbString:  vbuttons = v
        Case VarType(v) = vbEmpty:          vbuttons = vbOKOnly
        Case Else
            If IsArray(v) Then
                vbuttons = v
            ElseIf TypeName(v) = "Collection" Or TypeName(v) = "Dictionary" Then
                Set vbuttons = v
            End If
    End Select
End Property

Public Property Get MsgLabel( _
              Optional ByVal msg_section As Long) As TypeMsgLabel
' ------------------------------------------------------------------------------
' Transfers a section's message UDT stored as array back into the UDT.
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    If dctSectionsLabel Is Nothing Then
        MsgLabel.Text = vbNullString
    Else
        If dctSectionsLabel.Exists(msg_section) Then
            vArry = dctSectionsLabel(msg_section)
            MsgLabel.FontBold = vArry(0)
            MsgLabel.FontColor = vArry(1)
            MsgLabel.FontItalic = vArry(2)
            MsgLabel.FontName = vArry(3)
            MsgLabel.FontSize = vArry(4)
            MsgLabel.FontUnderline = vArry(5)
            MsgLabel.Monospaced = vArry(6)
            MsgLabel.Text = vArry(7)
        Else
            MsgLabel.Text = vbNullString
        End If
    End If
End Property

Public Property Let MsgLabel( _
              Optional ByVal msg_section As Long, _
                       ByRef msg_label As TypeMsgLabel)
' ------------------------------------------------------------------------------
' Transfers a message label UDT (msg_label) into an array and stores it in the
' Dictionary (dctSectionsLabel) with the section (msg_section) as the key.
' ------------------------------------------------------------------------------
    Dim vArry(7)    As Variant
    
    If dctSectionsLabel Is Nothing Then Set dctSectionsLabel = New Dictionary
    If Not dctSectionsLabel.Exists(msg_section) Then
        vArry(0) = msg_label.FontBold
        vArry(1) = msg_label.FontColor
        vArry(2) = msg_label.FontItalic
        vArry(3) = msg_label.FontName
        vArry(4) = msg_label.FontSize
        vArry(5) = msg_label.FontUnderline
        vArry(6) = msg_label.Monospaced
        vArry(7) = msg_label.Text
        dctSectionsLabel.Add msg_section, vArry
    End If
End Property

Public Property Get MsgMonoSpaced(Optional ByVal msg_section As Long) As Boolean
    Dim vArry() As Variant
    
    If dctSectionsText Is Nothing Then
        MsgMonoSpaced = False
    Else
        With dctSectionsText
            If .Exists(msg_section) Then
                vArry = .Item(msg_section)
                MsgMonoSpaced = vArry(6)
            Else
                MsgMonoSpaced = False
            End If
        End With
    End If
End Property

Public Property Get MsgText( _
             Optional ByVal msg_section As Long) As TypeMsgText
' ------------------------------------------------------------------------------
' Transferes message UDT stored as array back into the UDT.
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    If dctSectionsText Is Nothing Then
        MsgText.Text = vbNullString
    Else
        If dctSectionsText.Exists(msg_section) Then
            vArry = dctSectionsText(msg_section)
            MsgText.FontBold = vArry(0)
            MsgText.FontColor = vArry(1)
            MsgText.FontItalic = vArry(2)
            MsgText.FontName = vArry(3)
            MsgText.FontSize = vArry(4)
            MsgText.FontUnderline = vArry(5)
            MsgText.Monospaced = vArry(6)
            MsgText.Text = vArry(7)
        Else
            MsgText.Text = vbNullString
        End If
    End If

End Property

Public Property Let MsgText( _
             Optional ByVal msg_section As Long, _
                      ByRef msg_msg As TypeMsgText)
' ------------------------------------------------------------------------------
' Transfers a message UDT into an array and stores it in a Dictionary with the
' section number (msg_section) as the key.
' ------------------------------------------------------------------------------
    Dim vArry(7)    As Variant
    
    If dctSectionsText Is Nothing Then Set dctSectionsText = New Dictionary
    If Not dctSectionsText.Exists(msg_section) Then
        vArry(0) = msg_msg.FontBold
        vArry(1) = msg_msg.FontColor
        vArry(2) = msg_msg.FontItalic
        vArry(3) = msg_msg.FontName
        vArry(4) = msg_msg.FontSize
        vArry(5) = msg_msg.FontUnderline
        vArry(6) = msg_msg.Monospaced
        vArry(7) = msg_msg.Text
        dctSectionsText.Add msg_section, vArry
    End If
End Property

Public Property Get MsgTitle() As String
    MsgTitle = Me.Caption
End Property

Public Property Let MsgTitle(ByVal s As String)
    sTitle = s
    SetupTitle
End Property

Public Property Get NoOfDesignedMsgSects() As Long ' -----------------------
    NoOfDesignedMsgSects = 4                       ' Global definition !!!!!
End Property                                          ' -----------------------

Private Property Get PrcntgHeightBttnsArea() As Single
    PrcntgHeightBttnsArea = Round(DsgnBttnsArea.Height / (DsgnMsgArea.Height + DsgnBttnsArea.Height), 2)
End Property

Private Property Get PrcntgHeightMsgArea() As Single
    PrcntgHeightMsgArea = Round(DsgnMsgArea.Height / (DsgnMsgArea.Height + DsgnBttnsArea.Height), 2)
End Property

Public Property Let ReplyWithIndex(ByVal b As Boolean):                                 bReplyWithIndex = b:                                                        End Property

Public Property Get VmarginButtons() As Single:                                         VmarginButtons = siVmarginButtons:                                          End Property

Public Property Let VmarginButtons(ByVal si As Single):                                 siVmarginButtons = si:                                                      End Property

Public Property Get VmarginFrames() As Single:                                          VmarginFrames = siVmarginFrames:                                            End Property

Public Property Let VmarginFrames(ByVal si As Single):                                  siVmarginFrames = VgridPos(si):                                             End Property

Public Sub PositionMessageOnScreen( _
           Optional ByVal pos_top_left As Boolean = False)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    
    On Error Resume Next
        
    With Me
        .StartupPosition = sup_Manual
        If pos_top_left Then
            .Left = 5
            .top = 5
        Else
            .Left = (VirtualScreenWidthPts - .Width) / 2
            .top = (VirtualScreenHeightPts - .Height) / 4
        End If
    End With
    
    '~~ First make sure the bottom right fits,
    '~~ then check if the top-left is still on the screen (which gets priority).
    With Me
        If ((.Left + .Width) > (VirtualScreenLeftPts + VirtualScreenWidthPts)) Then .Left = ((VirtualScreenLeftPts + VirtualScreenWidthPts) - .Width)
        If ((.top + .Height) > (VirtualScreenTopPts + VirtualScreenHeightPts)) Then .top = ((VirtualScreenTopPts + VirtualScreenHeightPts) - .Height)
        If (.Left < VirtualScreenLeftPts) Then .Left = VirtualScreenLeftPts
        If (.top < VirtualScreenTopPts) Then .top = VirtualScreenTopPts
    End With
    
End Sub

Public Function AppErr(ByVal lNo As Long) As Long
' ------------------------------------------------------------------------------
' Converts a positive (i.e. an "application" error number into a negative number
' by adding vbObjectError. Converts a negative number back into a positive i.e.
' the original programmed application error number.
' Usage example:
'    Err.Raise mErH.AppErr(1), .... ' when an application error is detected
'    If Err.Number < 0 Then    ' when the error is displayed
'       MsgBox "Application error " & AppErr(Err.Number)
'    Else
'       MsgBox "VB Rutime Error " & Err.Number
'    End If
' ------------------------------------------------------------------------------
    AppErr = IIf(lNo < 0, AppErr = lNo - vbObjectError, AppErr = vbObjectError + lNo)
End Function

Private Sub ApplyHorizontalScrollBarButtonsIfRequired()
    Const PROC = "ApplyHorizontalScrollBarButtonsIfRequired"
    
    On Error GoTo eh
    Dim frBttnsArea      As MSForms.Frame: Set frBttnsArea = DsgnBttnsArea
    
    If frBttnsArea.Width > BttnsAreaWidthUsable Then
'        Debug_Sizes "Buttons area width exceeds maximum width specified:"
        ApplyScrollBarHorizontal scroll_frame:=frBttnsArea, new_width:=BttnsAreaWidthUsable
        bHscrollbarBttnsArea = True
        FormWidth = siMaxMsgWidthPts
        CenterHorizontal frBttnsArea
'        Debug_Sizes "Buttons area width decremented:"
    End If

xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub ApplyScrollBarHorizontal( _
                               ByVal scroll_frame As MSForms.Frame, _
                               ByVal new_width As Single)
' ------------------------------------------------------------------------------
' Apply a horizontal scrollbar either for the message area when a monospaced
' message section exceeds the specified maximum message form width or for the
' buttons area when one of the button rows exceeds the specified maximum message
' form width.
' ------------------------------------------------------------------------------
    Dim WidthOfScrollBar As Single
    
    With scroll_frame
        WidthOfScrollBar = .Width + 1
        .Width = new_width
        .Height = .Height + SPACE_VERTICAL_SCROLLBAR
        Select Case .ScrollBars
            Case fmScrollBarsBoth
                .KeepScrollBarsVisible = fmScrollBarsBoth
            Case fmScrollBarsHorizontal
                .ScrollWidth = WidthOfScrollBar
                .Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
                .KeepScrollBarsVisible = fmScrollBarsBoth
            Case fmScrollBarsNone, fmScrollBarsVertical
                .ScrollBars = fmScrollBarsHorizontal
                .ScrollWidth = WidthOfScrollBar
                .Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
                .KeepScrollBarsVisible = fmScrollBarsVertical
        End Select
    End With

End Sub

Private Sub ApplyScrollBarVertical( _
                             ByVal scroll_frame As MSForms.Frame, _
                             ByVal new_height As Single)
' ------------------------------------------------------------------------------
' Apply a vertical scrollbar for the frame (scroll_frame) and reduce the
' frame's height to the (new_height) while scrollbar's height will be the the
' original frame's height becomes the height of the scrollbar.
' ------------------------------------------------------------------------------
        
    Dim HeightOfScrollBar As Single
        
    With scroll_frame
        HeightOfScrollBar = .Height + SPACE_VERTICAL_SCROLLBAR
        .Height = new_height
        Select Case .ScrollBars
            Case fmScrollBarsHorizontal
                .ScrollBars = fmScrollBarsBoth
                .ScrollHeight = HeightOfScrollBar
                .KeepScrollBarsVisible = fmScrollBarsBoth
            Case fmScrollBarsNone
                .ScrollBars = fmScrollBarsVertical
                .ScrollHeight = HeightOfScrollBar
                .KeepScrollBarsVisible = fmScrollBarsVertical
            Case fmScrollBarsVertical
                .ScrollHeight = HeightOfScrollBar
                .KeepScrollBarsVisible = fmScrollBarsVertical
                .Scroll yAction:=fmScrollActionEnd
        End Select
    End With
    
End Sub

Private Sub ButtonClicked(ByVal cmb As MSForms.CommandButton)
' ------------------------------------------------------------------------------
' Return the value of the clicked reply button (button). When there is only one
' applied reply button the form is unloaded with the click of it. Otherwise the
' form is just hidden waiting for the caller to obtain the return value or
' index which then unloads the form.
' ------------------------------------------------------------------------------
    On Error Resume Next
    If bReplyWithIndex Then
        vReplyValue = ClickedButtonIndex(cmb)
        mMsg.RepliedWith = ClickedButtonIndex(cmb)
    Else
        vReplyValue = AppliedButtonRetVal(cmb)  ' global variable of calling module mMsg
        mMsg.RepliedWith = AppliedButtonRetVal(cmb)  ' global variable of calling module mMsg
    End If
    
    DisplayDone = True ' in case the form has been displayed modeless this will indicate the end of the wait loop
    Unload Me
    
End Sub

Private Sub CenterHorizontal(ByVal fr_center As MSForms.Frame, _
          Optional ByVal fr_within As MSForms.Frame = Nothing)
' ------------------------------------------------------------------------------
' Center the frame (fr) horizontally within the frame (frin)
' which defaults to the UserForm when not provided.
' ------------------------------------------------------------------------------
    
    If fr_within Is Nothing _
    Then fr_center.Left = (Me.InsideWidth - fr_center.Width) / 2 _
    Else fr_center.Left = (fr_within.Width - fr_center.Width) / 2
    
End Sub

' ------------------------------------------------------------
' The reply button click event is the only code using the
' control's name - which unfortunately this cannot be avioded.
' ------------------------------------------------------------
Private Sub cmb11_Click():  ButtonClicked Me.cmb11:   End Sub

Private Sub cmb12_Click():  ButtonClicked Me.cmb12:   End Sub

Private Sub cmb13_Click():  ButtonClicked Me.cmb13:   End Sub

Private Sub cmb14_Click():  ButtonClicked Me.cmb14:   End Sub

Private Sub cmb15_Click():  ButtonClicked Me.cmb15:   End Sub

Private Sub cmb16_Click():  ButtonClicked Me.cmb16:   End Sub

Private Sub cmb17_Click():  ButtonClicked Me.cmb17:   End Sub

Private Sub cmb21_Click():  ButtonClicked Me.cmb21:   End Sub

Private Sub cmb22_Click():  ButtonClicked Me.cmb22:   End Sub

Private Sub cmb23_Click():  ButtonClicked Me.cmb23:   End Sub

Private Sub cmb24_Click():  ButtonClicked Me.cmb24:   End Sub

Private Sub cmb25_Click():  ButtonClicked Me.cmb25:   End Sub

Private Sub cmb26_Click():  ButtonClicked Me.cmb26:   End Sub

Private Sub cmb27_Click():  ButtonClicked Me.cmb27:   End Sub

Private Sub cmb31_Click():  ButtonClicked Me.cmb31:   End Sub

Private Sub cmb32_Click():  ButtonClicked Me.cmb32:   End Sub

Private Sub cmb33_Click():  ButtonClicked Me.cmb33:   End Sub

Private Sub cmb34_Click():  ButtonClicked Me.cmb34:   End Sub

Private Sub cmb35_Click():  ButtonClicked Me.cmb35:   End Sub

Private Sub cmb36_Click():  ButtonClicked Me.cmb36:   End Sub

Private Sub cmb37_Click():  ButtonClicked Me.cmb37:   End Sub

Private Sub cmb41_Click():  ButtonClicked Me.cmb41:   End Sub

Private Sub cmb42_Click():  ButtonClicked Me.cmb42:   End Sub

Private Sub cmb43_Click():  ButtonClicked Me.cmb43:   End Sub

Private Sub cmb44_Click():  ButtonClicked Me.cmb44:   End Sub

Private Sub cmb45_Click():  ButtonClicked Me.cmb45:   End Sub

Private Sub cmb46_Click():  ButtonClicked Me.cmb46:   End Sub

Private Sub cmb47_Click():  ButtonClicked Me.cmb47:   End Sub

Private Sub cmb51_Click():  ButtonClicked Me.cmb51:   End Sub

Private Sub cmb52_Click():  ButtonClicked Me.cmb52:   End Sub

Private Sub cmb53_Click():  ButtonClicked Me.cmb53:   End Sub

Private Sub cmb54_Click():  ButtonClicked Me.cmb54:   End Sub

Private Sub cmb55_Click():  ButtonClicked Me.cmb55:   End Sub

Private Sub cmb56_Click():  ButtonClicked Me.cmb56:   End Sub

Private Sub cmb57_Click():  ButtonClicked Me.cmb57:   End Sub

Private Sub cmb61_Click():  ButtonClicked Me.cmb61:   End Sub

Private Sub cmb62_Click():  ButtonClicked Me.cmb62:   End Sub

Private Sub cmb63_Click():  ButtonClicked Me.cmb63:   End Sub

Private Sub cmb64_Click():  ButtonClicked Me.cmb64:   End Sub

Private Sub cmb65_Click():  ButtonClicked Me.cmb65:   End Sub

Private Sub cmb66_Click():  ButtonClicked Me.cmb66:   End Sub

Private Sub cmb67_Click():  ButtonClicked Me.cmb67:   End Sub

Private Sub cmb71_Click():  ButtonClicked Me.cmb71:   End Sub

Private Sub cmb72_Click():  ButtonClicked Me.cmb72:   End Sub

Private Sub cmb73_Click():  ButtonClicked Me.cmb73:   End Sub

Private Sub cmb74_Click():  ButtonClicked Me.cmb74:   End Sub

Private Sub cmb75_Click():  ButtonClicked Me.cmb75:   End Sub

Private Sub cmb76_Click():  ButtonClicked Me.cmb76:   End Sub

Private Sub cmb77_Click():  ButtonClicked Me.cmb77:   End Sub

Private Sub Collect(ByRef cllct_into As Variant, _
                    ByVal cllct_with_parent As Variant, _
                    ByVal cllct_cntrl_type As String, _
                    ByVal cllct_set_height As Single, _
                    ByVal cllct_set_width As Single, _
           Optional ByVal cllct_set_visible As Boolean = False)
' ------------------------------------------------------------------------------
' Setup of a Collection (cllct_into) with all type (cllct_cntrl_type) controls
' with a parent (cllct_with_parent) as Collection (cllct_into) by assigning the
' an initial height (cllct_set_height) and width (cllct_set_width).
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim ctl As MSForms.Control
    Dim v   As Variant
     
    Set cllct_into = New Collection

    Select Case TypeName(cllct_with_parent)
        Case "Collection"
            '~~ Parent is each frame in the collection
            For Each v In cllct_with_parent
                For Each ctl In Me.Controls
                    If TypeName(ctl) = cllct_cntrl_type And ctl.Parent Is v Then
                        With ctl
                            .Visible = cllct_set_visible
                            .Height = cllct_set_height
                            .Width = cllct_set_width
                        End With
                        cllct_into.Add ctl
                    End If
               Next ctl
            Next v
        Case Else
            For Each ctl In Me.Controls
                If TypeName(ctl) = cllct_cntrl_type And ctl.Parent Is cllct_with_parent Then
                    With ctl
                        .Visible = cllct_set_visible
                        .Height = cllct_set_height
                        .Width = cllct_set_width
                    End With
                    Select Case TypeName(cllct_into)
                        Case "Collection"
                            cllct_into.Add ctl
                        Case Else
                            Set cllct_into = ctl
                    End Select
                End If
            Next ctl
    End Select

xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub CollectDesignControls()
' ----------------------------------------------------------------------
' Collects all designed controls without concidering any control's name.
' ----------------------------------------------------------------------
    Const PROC = "CollectDesignControls"
    
    On Error GoTo eh
    Dim v As Variant

    Collect cllct_into:=cllDsgnAreas _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=Me _
          , cllct_set_height:=10 _
          , cllct_set_width:=Me.Width - siHmarginFrames
    
    Collect cllct_into:=cllDsgnSections _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=DsgnMsgArea _
          , cllct_set_height:=50 _
          , cllct_set_width:=DsgnMsgArea.Width - siHmarginFrames
    
    Collect cllct_into:=cllDsgnSectionsLabel _
          , cllct_cntrl_type:="Label" _
          , cllct_with_parent:=cllDsgnSections _
          , cllct_set_height:=15 _
          , cllct_set_width:=DsgnMsgArea.Width - (siHmarginFrames * 2)
    
    Collect cllct_into:=cllDsgnSectionsTextFrame _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=cllDsgnSections _
          , cllct_set_height:=20 _
          , cllct_set_width:=DsgnMsgArea.Width - (siHmarginFrames * 2)
    
    Collect cllct_into:=cllDsgnSectionsText _
          , cllct_cntrl_type:="TextBox" _
          , cllct_with_parent:=cllDsgnSectionsTextFrame _
          , cllct_set_height:=20 _
          , cllct_set_width:=DsgnMsgArea.Width - (siHmarginFrames * 3)
    
    Collect cllct_into:=cllDsgnBttnsFrame _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=DsgnBttnsArea _
          , cllct_set_height:=10 _
          , cllct_set_width:=10 _
          , cllct_set_visible:=True ' minimum is one button
    
    Collect cllct_into:=cllDsgnBttnRows _
          , cllct_cntrl_type:="Frame" _
          , cllct_with_parent:=cllDsgnBttnsFrame _
          , cllct_set_height:=10 _
          , cllct_set_width:=10 _
          , cllct_set_visible:=False ' minimum is one button
        
    Set cllDsgnBttns = New Collection
    For Each v In cllDsgnBttnRows
        Collect cllct_into:=cllDsgnRowBttns _
              , cllct_cntrl_type:="CommandButton" _
              , cllct_with_parent:=v _
              , cllct_set_height:=10 _
              , cllct_set_width:=siMinButtonWidth
        cllDsgnBttns.Add cllDsgnRowBttns
    Next v
    
    ProvideDictionary dctAppliedControls ' provides a clean or new dictionary for collection applied controls
    ProvideDictionary AppliedBttns
    ProvideDictionary AppliedBttnsRetVal
    ProvideDictionary AppliedBttnRows

xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub
 
Private Sub ConvertPixelsToPoints(ByVal x_dpi As Single, ByVal y_dpi As Single, _
                                  ByRef x_pts As Single, ByRef y_pts As Single)
' ------------------------------------------------------------------------------
' Returns pixels (device dependent) to points (used by Excel).
' ------------------------------------------------------------------------------
    
    Dim hDC            As Long
    Dim RetVal         As Long
    Dim PixelsPerInchX As Long
    Dim PixelsPerInchY As Long
 
    On Error Resume Next
    hDC = GetDC(0)
    PixelsPerInchX = GetDeviceCaps(hDC, LOGPIXELSX)
    PixelsPerInchY = GetDeviceCaps(hDC, LOGPIXELSY)
    RetVal = ReleaseDC(0, hDC)
    x_pts = x_dpi * TWIPSPERINCH / 20 / PixelsPerInchX
    y_pts = y_dpi * TWIPSPERINCH / 20 / PixelsPerInchY

End Sub

Private Sub DisplayFramesWithCaptions( _
                        Optional ByVal b As Boolean = True)
' ---------------------------------------------------------
' When False (the default) captions are removed from all
' frames Else they remain visible for testing purpose
' ---------------------------------------------------------
            
    Dim ctl As MSForms.Control
       
    If Not b Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.Caption = vbNullString
            End If
        Next ctl
    End If

End Sub

Private Function DsgnRowBttns(ByVal buttonrow As Long) As Collection
' --------------------------------------------------------------------
' Return a collection of applied/use/visible buttons in row buttonrow.
' --------------------------------------------------------------------
    Set DsgnRowBttns = cllDsgnBttns(buttonrow)
End Function

Private Sub ErrMsg( _
             ByVal err_source As String, _
    Optional ByVal err_no As Long = 0, _
    Optional ByVal err_dscrptn As String = vbNullString)
' ------------------------------------------------------
' This Common Component does not have its own error
' handling. Instead it passes on any error to the
' caller's error handling.
' ------------------------------------------------------
    
    If err_no = 0 Then err_no = Err.Number
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    Debug.Print "Error in: " & err_source & ": Error = " & err_no & " " & err_dscrptn
'    Err.Raise Number:=err_no, Source:=err_source, Description:=err_dscrptn

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "fMsg." & sProc
End Function

Private Sub GetScreenMetrics(ByRef left_pts As Single, _
                             ByRef top_pts As Single, _
                             ByRef width_pts As Single, _
                             ByRef height_pts As Single)
' ------------------------------------------------------------
' Get coordinates of top-left corner and size of entire screen
' (stretched over all monitors) and convert to Points.
' ------------------------------------------------------------
    
    ConvertPixelsToPoints x_dpi:=GetSystemMetrics32(SM_XVIRTUALSCREEN), x_pts:=left_pts, _
                          y_dpi:=GetSystemMetrics32(SM_YVIRTUALSCREEN), y_pts:=top_pts
                          
    ConvertPixelsToPoints x_dpi:=GetSystemMetrics32(SM_CXVIRTUALSCREEN), x_pts:=width_pts, _
                          y_dpi:=GetSystemMetrics32(SM_CYVIRTUALSCREEN), y_pts:=height_pts

End Sub

Private Function Max(ParamArray va() As Variant) As Variant
' ---------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ---------------------------------------------------------
    Dim v   As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function Min(ParamArray va() As Variant) As Variant
' ------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ------------------------------------------------------
    Dim v   As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

'Private Sub ProvideCollection(ByRef cll As Collection)
'' ----------------------------------------------------
'' Provides a clean/new Collection.
'' ----------------------------------------------------
'    If Not cll Is Nothing Then Set cll = Nothing
'    Set cll = New Collection
'End Sub

Private Sub ProvideDictionary(ByRef dct As Dictionary)
' ----------------------------------------------------
' Provides a clean or new Dictionary.
' ----------------------------------------------------
    If Not dct Is Nothing Then dct.RemoveAll Else Set dct = New Dictionary
End Sub

Private Sub ReduceAreasHeightAndApplyVerticalScrollBar(ByVal total_exceeding_height As Single)
' -------------------------------------------------------------------
' Reduce the final form height to the maximum height specified by
' reducing either the buttons or the message area by the exceeding
' height and apply a vertcal scrollbar. The area which uses 60%
' or more of the overall height is the one being reduced. Otherwise
' both are reduced proportionally.
' -------------------------------------------------------------------

    Dim frMsgArea               As MSForms.Frame:   Set frMsgArea = DsgnMsgArea
    Dim frBttnsArea             As MSForms.Frame:   Set frBttnsArea = DsgnBttnsArea
    Dim siAreasExceedingHeight  As Single
    
    With Me
        '~~ Reduce the height to the max height specified
        siAreasExceedingHeight = .Height - siMaxMsgHeightPts
        .Height = siMaxMsgHeightPts
        
        If PrcntgHeightMsgArea >= 0.6 Then
            '~~ When the message area requires 60% or more of the total height only this frame
            '~~ will be reduced and applied with a vertical scrollbar.
            ApplyScrollBarVertical scroll_frame:=frMsgArea _
                                 , new_height:=frMsgArea.Height - total_exceeding_height
            HeightDecrementMsgArea = True
            
        ElseIf PrcntgHeightBttnsArea >= 0.6 Then
            '~~ When the buttons area requires 60% or more it will be reduced and applied with a vertical scrollbar.
            ApplyScrollBarVertical scroll_frame:=frBttnsArea _
                                 , new_height:=frBttnsArea.Height - total_exceeding_height
            HeightDecrementBttnsArea = True

        Else
            '~~ When one area of the two requires less than 60% of the total areas heigth
            '~~ both will be reduced in the height and get a vertical scrollbar.
            ApplyScrollBarVertical scroll_frame:=frMsgArea _
                                 , new_height:=frMsgArea.Height * PrcntgHeightMsgArea
            HeightDecrementMsgArea = True
            ApplyScrollBarVertical scroll_frame:=frBttnsArea _
                                 , new_height:=frBttnsArea.Height * PrcntgHeightBttnsArea
            HeightDecrementBttnsArea = True
        End If
    End With
    If DsgnMsgArea.Width > Me.InsideWidth Then
        FormWidth = DsgnMsgArea.Width
    End If
    
End Sub

Private Sub ResizeAndRepositionAreas()

    Dim v       As Variant
    Dim siTop   As Single
    
    siTop = siVmarginFrames
    For Each v In cllDsgnAreas
        With v
            If IsApplied(v) Then
                .Visible = True
                .top = siTop
                siTop = VgridPos(.top + .Height + SPACE_VERTICAL_AREAS)
            End If
        End With
    Next v
    Me.Height = VgridPos(siTop + (SPACE_VERTICAL_AREAS * 3))
    
End Sub

Private Sub ResizeAndRepositionButtonRows()
' ------------------------------------------------------------------------------
' Adjust all applied/visible button rows height to the maximum buttons height
' and the row frames width to the number of displayed buttons - by keeping a
' record of the resulting maximum row's width.
' ------------------------------------------------------------------------------
    Const PROC = "ResizeAndRepositionButtonRows"
    
    On Error GoTo eh
    Dim BttnRowFrames       As Collection:      Set BttnRowFrames = DsgnBttnRows
    Dim BttnRowFrame        As MSForms.Frame
    Dim siTop               As Single
    Dim vButton             As Variant
    Dim lRow                As Long
    Dim AppliedRowFrames    As New Dictionary
    Dim v                   As Variant
    Dim lButtons            As Long
    Dim siHeight            As Single
    Dim BttnsFrameWidth  As Single
    
    '~~ Collect the applied/visible rows as key and the
    '~~ number of applied/visible buttons therein as item
    For lRow = 1 To BttnRowFrames.Count
        Set BttnRowFrame = BttnRowFrames(lRow)
        If IsApplied(BttnRowFrame) Then
            With BttnRowFrame
                .Visible = True
                lButtons = 0
                For Each vButton In DsgnRowBttns(lRow)
                    If IsApplied(vButton) Then lButtons = lButtons + 1
                Next vButton
                AppliedRowFrames.Add BttnRowFrame, lButtons
            End With
        End If
    Next lRow
    
    '~~ Adjust button row's width and height
    siHeight = AppliedButtonRowHeight
    siTop = siVmarginFrames
    For Each v In AppliedRowFrames
        Set BttnRowFrame = v
        lButtons = AppliedRowFrames(v)
        With BttnRowFrame
            .top = siTop
            .Height = siHeight
            '~~ Provide some extra space for the button's design
            BttnsFrameWidth = CInt((siMaxButtonWidth * lButtons) _
                           + (siHmarginButtons * (lButtons - 1)) _
                           + (siHmarginFrames * 2)) - siHmarginButtons + 7
            .Width = BttnsFrameWidth
            siBttnsFrameMaxWidth = Max(.Width, siBttnsFrameMaxWidth)
            siTop = .top + .Height + siVmarginButtons
        End With
    Next v
    Set AppliedRowFrames = Nothing
    siBttnsFrameMaxWidth = siBttnsFrameMaxWidth + 5 ' add design space

xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub ResizeAndRepositionButtons()
' ------------------------------------------------------------------------------
' Unify all applied/visible button's size by assigning the maximum width and
' height provided with their setup, and adjust their resulting left position.
' ------------------------------------------------------------------------------
    Const PROC = "ResizeAndRepositionButtons"
    
    On Error GoTo eh
    Dim cllButtonRows   As Collection:      Set cllButtonRows = DsgnBttnRows
    Dim siLeft          As Single
    Dim frRow           As MSForms.Frame
    Dim vButton         As Variant
    Dim lRow            As Long
    Dim lButton         As Long
    
    For lRow = 1 To cllButtonRows.Count
        siLeft = siHmarginFrames
        Set frRow = cllButtonRows(lRow)
        If IsApplied(frRow) Then
            For Each vButton In DsgnRowBttns(lRow)
                If IsApplied(vButton) Then
                    lButton = lButton + 1
                    With vButton
                        .Visible = True
                        .Left = siLeft
                        .Width = siMaxButtonWidth
                        .Height = siMaxButtonHeight
                        .top = siVmarginFrames
                        siLeft = .Left + .Width + siHmarginButtons
                        If IsNumeric(vDefaultBttn) Then
                            If lButton = vDefaultBttn Then .Default = True
                        Else
                            If .Caption = vDefaultBttn Then .Default = True
                        End If
                    End With
                End If
            Next vButton
        End If
    Next lRow
        
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub ResizeAndRepositionBttnsArea()
' ----------------------------------------------------
' Adjust buttons frame to max button row width and the
' surrounding area's width and heigth is adjusted
' ----------------------------------------------------
    Const PROC = "ResizeAndRepositionBttnsArea"
    
    On Error GoTo eh
    Dim frBttnsArea     As MSForms.Frame:   Set frBttnsArea = DsgnBttnsArea
    Dim frBttnsFrame    As MSForms.Frame:   Set frBttnsFrame = DsgnBttnsFrame
        
    If IsApplied(frBttnsFrame) Then
        With frBttnsArea
            .Left = SPACE_HORIZONTAL_BTTN_AREA
            .Visible = True
            Select Case .ScrollBars
                Case fmScrollBarsBoth
                    .Width = siBttnsFrameMaxWidth + (siHmarginFrames * 2) + SPACE_HORIZONTAL_SCROLLBAR + 2 ' space reserved or used
                    frBttnsFrame.Left = 0
                Case fmScrollBarsHorizontal
                    .Width = siBttnsFrameMaxWidth + (siHmarginFrames * 2) + 2
                Case fmScrollBarsNone
                    .Height = frBttnsFrame.Height + (siVmarginFrames * 2)
                    .Width = siBttnsFrameMaxWidth + (siHmarginFrames * 2) + 2
                Case fmScrollBarsVertical
                    .Width = siBttnsFrameMaxWidth + (siHmarginFrames * 2) + SPACE_HORIZONTAL_SCROLLBAR + 2 ' space reserved or used
            End Select
            
            FormWidth = (.Width + siHmarginFrames * 2) + (SPACE_HORIZONTAL_BTTN_AREA * 2)
            
            If .ScrollBars = fmScrollBarsNone _
            Then CenterHorizontal fr_center:=frBttnsFrame, fr_within:=frBttnsArea
            Me.Height = .top + .Height + SPACE_VERTICAL_BOTTOM
        End With
    
        CenterHorizontal fr_center:=frBttnsArea
    End If
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub ResizeAndRepositionBttnsFrame()
' ------------------------------------------------------------------------------
' Adjust the frame around all button row frames to the maximum widht calculated
' by the adjustment of each of the rows frame.
' ------------------------------------------------------------------------------
    Const PROC = "ResizeAndRepositionBttnsFrame"
    
    On Error GoTo eh
    Dim frBttnsFrame  As MSForms.Frame: Set frBttnsFrame = DsgnBttnsFrame
    Dim v   As Variant
    
    If IsApplied(frBttnsFrame) Then
        With frBttnsFrame
            .Visible = True
            .top = siVmarginFrames
            .Width = siBttnsFrameMaxWidth + (siHmarginFrames * 2) ' Left and right margin added
            .Height = BttnsFrameHeight
            If bVscrollbarBttnsArea _
            Then .Left = siHmarginFrames _
            Else .Left = siHmarginFrames + (SPACE_HORIZONTAL_SCROLLBAR / 2)
        End With
    End If
    '~~ Center all button rows
    For Each v In DsgnBttnRows
        If IsApplied(v) Then CenterHorizontal fr_center:=v, fr_within:=frBttnsFrame
    Next v
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub ResizeAndRepositionMsgArea()
' --------------------------------------------------------
' Re-position all applied/used message sections vertically
' and adjust the Message Area height accordingly.
' --------------------------------------------------------
    Const PROC = "ResizeAndRepositionMsgArea"
    
    On Error GoTo eh
    Dim frArea      As MSForms.Frame: Set frArea = DsgnMsgArea
    Dim siTop       As Single
            
    If IsApplied(frArea) Then
        siTop = siVmarginFrames
        Me.Height = Max(Me.Height, frArea.top + frArea.Height + (SPACE_VERTICAL_AREAS * 4))
    End If
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub ResizeAndRepositionMsgSects()
' -----------------------------------------------------------
' Assign all displayed message sections the required height
' and finally adjust the message area's height.
' -----------------------------------------------------------
    Const PROC = "ResizeAndRepositionMsgSects"
    
    On Error GoTo eh
    Dim frSection       As MSForms.Frame
    Dim i               As Long
    Dim la              As MSForms.Label
    Dim frText          As MSForms.Frame
    Dim tb              As MSForms.TextBox
    Dim siTop           As Single
    Dim siTopSection    As Single
    
    siTopSection = 6
    For i = 1 To cllDsgnSections.Count
        siTop = 0
        If IsApplied(DsgnSection(i)) Then
            DsgnMsgArea.Visible = True
            Set frSection = DsgnSection(i)
            frSection.Visible = True
            Set la = DsgnSectionLabel(i)
            Set frText = DsgnSectionTextFrame(i)
            Set tb = DsgnSectionText(i)
            frSection.Width = DsgnMsgArea.Width - siHmarginFrames
            
            If IsApplied(la) Then
                With la
                    .Visible = True
                    .top = siTop
                    .Width = frSection.Width - siHmarginFrames
                    siTop = VgridPos(.top + .Height)
                End With
            End If
            
            If IsApplied(tb) Then
                With tb
                    .Visible = True
                    .top = siVmarginFrames
                End With
                With frText
                    .Visible = True
                    .top = siTop
'                    .Height = tb.Height + (siVmarginFrames * 2)
                    siTop = .top + .Height + siVmarginFrames
                    If .ScrollBars = fmScrollBarsBoth Or frText.ScrollBars = fmScrollBarsHorizontal Then
                        .Height = tb.top + tb.Height + SPACE_VERTICAL_SCROLLBAR + siVmarginFrames
                    Else
                        .Height = tb.top + tb.Height + siVmarginFrames
                    End If
                End With
            End If
                
            If IsApplied(frSection) Then
                With frSection
                    .top = siTopSection
                    .Visible = True
                    .Height = frText.top + frText.Height + siVmarginFrames
                    siTopSection = VgridPos(.top + .Height + siVmarginFrames + SPACE_VERTICAL_SECTIONS)
                End With
            End If
                
            Select Case DsgnMsgArea.ScrollBars
                Case fmScrollBarsBoth, fmScrollBarsVertical:    frSection.Left = siHmarginFrames
                Case fmScrollBarsHorizontal, fmScrollBarsNone:  frSection.Left = siHmarginFrames + (SPACE_VERTICAL_SCROLLBAR / 2)
            End Select
        End If
    Next i
    
    If Not frSection Is Nothing Then DsgnMsgArea.Height = frSection.top + frSection.Height + siVmarginFrames
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
End Sub

Public Sub Progress( _
              ByVal prgrs_text As String, _
     Optional ByVal prgrs_section As Long = 1, _
     Optional ByVal prgrs_append As Boolean = True)
' ------------------------------------------------------------------------------
' Replaces the section's (prgrs_section) text with (prgrs_text) or appends it
' when (prgrs_append) = True.
' ------------------------------------------------------------------------------
    
    If MsgText(prgrs_section).Monospaced _
    Then SetupMsgSectMonoSpaced msg_section:=prgrs_section _
                              , msg_append:=prgrs_append _
                              , msg_text:=prgrs_text _
    Else SetupMsgSectPropSpaced msg_section:=prgrs_section _
                              , msg_append:=prgrs_append _
                              , msg_text:=prgrs_text

'    With DsgnSectionText(prgrs_section)
'        .SetFocus
'        If prgrs_append _
'        Then .Value = .Value & vbLf & prgrs_text _
'        Else .Value = prgrs_text
'        DoEvents
'    End With
    
    ResizeAndRepositionMsgSects
    ResizeAndRepositionMsgArea
    ResizeAndRepositionAreas
    DoEvents

    If Me.Height > siMaxMsgHeightPts Then
        '~~ The message form height exceeds the specified or the default message height.
        '~~ The message area or the buttons area or both will be reduced to meet the
        '~~ limit and a vertical scrollbar will be setup. When both areas are about the
        '~~ same height (neither is taller the than 60% of the total heigth, both will
        '~~ get a vertical scrollbar, else only the one which uses 60% or more of the height.
        ReduceAreasHeightAndApplyVerticalScrollBar Me.Height - siMaxMsgHeightPts
        bDoneHeightDecrement = True
    End If
    
    ResizeAndRepositionAreas

End Sub

Public Sub Setup()
    Const PROC = "Setup"
    
    On Error GoTo eh
    
    CollectDesignControls
       
    DisplayFramesWithCaptions bDsplyFrmsWthCptnTestOnly ' may be True for test purpose
    
    '~~ Start the setup as if there wouldn't be any message - which might be the case
    Me.StartupPosition = 2
    FormWidth = Min(Me.Width, MinMsgWidthPts) ' may have been enlarged for the title when already setup
    Me.Height = 200                             ' just to start with - specifically for test purpose
    PositionMessageOnScreen pos_top_left:=True  ' in case of test best pos to start with
    DsgnMsgArea.Visible = False
    DsgnBttnsArea.top = SPACE_VERTICAL_AREAS
    
    '~~ ----------------------------------------------------------------------------------------
    '~~ The  p r i m a r y  setup of the title, the message sections and the reply buttons
    '~~ returns their individual widths which determines the minimum required message form width
    '~~ This setup ends width the final message form width and all elements adjusted to it.
    '~~ ----------------------------------------------------------------------------------------
    '~~ Setup of the title - may determine the final message width
    If Not bDoneTitle Then SetupTitle
    
    '~~ Setup of any monospaced message sections - may determine the final message width
    SetupMsgSectsMonoSpaced
    If IsApplied(DsgnMsgArea) Then
        ResizeAndRepositionMsgSects
        ResizeAndRepositionMsgArea
    End If
    
    '~~ Setup the reply buttons - may determine the final message section
    SetupBttns vbuttons
    ResizeAndRepositionButtons
    ResizeAndRepositionButtonRows
    ResizeAndRepositionBttnsFrame
    ResizeAndRepositionBttnsArea
    ApplyHorizontalScrollBarButtonsIfRequired
    If bVscrollbarBttnsArea Or bHscrollbarBttnsArea Then ResizeAndRepositionBttnsArea

    ' -----------------------------------------------------------------------------------------
    ' At this point the form width is final. The message height only when there are just
    ' buttons but no message. The setup of proportional spaced message sections will
    ' determine the final message height - when exeeding the maximum with a vertical scrollbar.
    ' -----------------------------------------------------------------------------------------
'    DsgnMsgArea.Width = Me.InsideWidth - siHmarginFrames
    
    '~~ Setup proportional spaced message sections by the given message form width
    SetupMsgSectsPropSpaced
    
    If IsApplied(DsgnMsgArea) Then
        ResizeAndRepositionMsgSects
        ResizeAndRepositionMsgArea
    End If
            
    If Me.Height > siMaxMsgHeightPts Then
        '~~ The message form height exceeds the specified or the default message height.
        '~~ The message area or the buttons area or both will be reduced to meet the
        '~~ limit and a vertical scrollbar will be setup. When both areas are about the
        '~~ same height (neither is taller the than 60% of the total heigth, both will
        '~~ get a vertical scrollbar, else only the one which uses 60% or more of the height.
        ReduceAreasHeightAndApplyVerticalScrollBar Me.Height - siMaxMsgHeightPts
        bDoneHeightDecrement = True
    End If
    
    ResizeAndRepositionAreas

    PositionMessageOnScreen
    bDoneSetup = True ' To indicate for the Activate event that the setup had already be done beforehand
    
xt: Exit Sub

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub SetupButton(ByVal buttonrow As Long, _
                        ByVal buttonindex As Long, _
                        ByVal buttoncaption As String, _
                        ByVal buttonreturnvalue As Variant)
' -----------------------------------------------------------------
' Setup an applied reply buttonindex's (buttonindex) visibility and
' caption, calculate the maximum buttonindex width and height,
' keep a record of the setup reply buttonindex's return value.
' -----------------------------------------------------------------
    Const PROC = "SetupButton"
    
    On Error GoTo eh
    Dim cmb As MSForms.CommandButton:   Set cmb = DsgnBttn(buttonrow, buttonindex)
    
    With cmb
        .Visible = True
        .AutoSize = True
        .WordWrap = False ' the longest line determines the buttonindex's width
        .Caption = buttoncaption
        .AutoSize = False
        .Height = .Height + 1 ' safety margin to ensure proper multilin caption display
        siMaxButtonHeight = Max(siMaxButtonHeight, .Height)
        siMaxButtonWidth = Max(siMaxButtonWidth, .Width, siMinButtonWidth)
    End With
    AppliedBttns.Add cmb, buttonrow
    AppliedButtonRetVal(cmb) = buttonreturnvalue ' keep record of the setup buttonindex's reply value
    AppliedControls = cmb
    AppliedControls = DsgnBttnRow(buttonrow)
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
End Sub

Private Sub SetupBttns(ByVal vbuttons As Variant)
' --------------------------------------------------------------------------------------
' Setup and position the applied reply buttons and calculate the max reply button width.
' Note: When the provided vButtons argument is a string it wil be converted into a
'       collection and the procedure is performed recursively with it.
' --------------------------------------------------------------------------------------
    Const PROC = "SetupBttns"
    
    On Error GoTo eh
    Dim frBttnsArea As MSForms.Frame:   Set frBttnsArea = DsgnBttnsArea
    
    AppliedControls = frBttnsArea
    AppliedControls = DsgnBttnsFrame
    lSetupRows = 1
    
    '~~ Setup all reply button by calculatig their maximum width and height
    Select Case TypeName(vbuttons)
        Case "Long":        SetupBttnsFromValue vbuttons ' buttons are specified by one single VBA.MsgBox button value only
        Case "String":      SetupBttnsFromString vbuttons
        Case "Collection":  SetupBttnsFromCollection vbuttons
        Case "Dictionary":  SetupBttnsFromCollection vbuttons
        Case Else
            '~~ Because vbuttons is not provided by a known/accepted format
            '~~ the message will be setup with an Ok only button", vbExclamation
            SetupBttns vbOKOnly
    End Select
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
End Sub

Private Sub SetupBttnsFromCollection(ByVal cllButtons As Collection)
' ---------------------------------------------------------------------
' Setup the reply buttons based on the comma delimited string of button
' captions and row breaks indicated by a vbLf, vbCr, or vbCrLf.
' ---------------------------------------------------------------------
    Const PROC = "SetupBttnsFromCollection"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim BttnsAreaFrame  As MSForms.Frame
    Dim BttnsFrame      As MSForms.Frame
    Dim BttnsRowFrame   As MSForms.Frame
    Dim Bttns
    Dim Bttn            As MSForms.CommandButton
    
    lSetupRows = 1
    lSetupRowButtons = 0
    Set BttnsAreaFrame = DsgnBttnsArea
    Set BttnsFrame = DsgnBttnsFrame
    Set BttnsRowFrame = DsgnBttnRow(1)
    Set Bttn = DsgnBttn(1, 1)
    
    Me.Height = 100 ' just to start with
    BttnsAreaFrame.top = SPACE_VERTICAL_AREAS
    BttnsAreaFrame.Visible = True
    BttnsFrame.top = BttnsAreaFrame.top
    BttnsFrame.Visible = True
    BttnsRowFrame.top = BttnsFrame.top
    BttnsRowFrame.Visible = True
    Bttn.top = BttnsRowFrame.top
    Bttn.Visible = True
    Bttn.Width = DEFAULT_BTTN_MIN_WIDTH
    
    For Each v In cllButtons
        Select Case v
            Case vbOKOnly
                SetupBttnsFromValue v
            Case vbOKCancel, vbYesNo, vbRetryCancel
                SetupBttnsFromValue v
            Case vbYesNoCancel, vbAbortRetryIgnore
                SetupBttnsFromValue v
            Case Else
                If v <> vbNullString Then
                    If v = vbLf Or v = vbCr Or v = vbCrLf Then
                        '~~ prepare for the next row
                        If lSetupRows <= 7 Then ' ignore exceeding rows
                            If Not AppliedBttnRows.Exists(DsgnBttnRow(lSetupRows)) Then AppliedBttnRows.Add DsgnBttnRow(lSetupRows), lSetupRows
                            AppliedControls = DsgnBttnRow(lSetupRows)
                            lSetupRows = lSetupRows + 1
                            lSetupRowButtons = 0
                        Else
                            MsgBox "Setup of button row " & lSetupRows & " ignored! The maximum applicable rows is 7."
                        End If
                    Else
                        lSetupRowButtons = lSetupRowButtons + 1
                        If lSetupRowButtons <= 7 Then
                            DsgnBttnRow(lSetupRows).Visible = True
                            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:=v, buttonreturnvalue:=v
                        Else
                            MsgBox "The setup of a button " & lSetupRowButtons & " in row " & lSetupRows & " is ignored! The maximum applicable buttons per row is 7."
                        End If
                    End If
                End If
        End Select
    Next v
    If lSetupRows <= 7 Then
        If Not AppliedBttnRows.Exists(DsgnBttnRow(lSetupRows)) Then AppliedBttnRows.Add DsgnBttnRow(lSetupRows), lSetupRows
        AppliedControls = DsgnBttnRow(lSetupRows)
    End If
    DsgnBttnsArea.Visible = True
    
xt: Exit Sub
eh: ErrMsg ErrSrc(PROC)
End Sub

Private Sub SetupBttnsFromString(ByVal buttons_string As String)
    
    Dim cll As New Collection
    Dim v   As Variant
    
    For Each v In Split(buttons_string, ",")
        cll.Add v
    Next v
    SetupBttns cll
    
End Sub

Private Sub SetupBttnsFromValue(ByVal lButtons As Long)
' -------------------------------------------------------
' Setup a row of standard VB MsgBox reply command buttons
' -------------------------------------------------------
    Const PROC = "SetupBttnsFromValue"
    
    On Error GoTo eh
    
    Select Case lButtons
        Case vbOKOnly
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Ok", buttonreturnvalue:=vbOK
        Case vbOKCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Ok", buttonreturnvalue:=vbOK
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbYesNo
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Yes", buttonreturnvalue:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="No", buttonreturnvalue:=vbNo
        Case vbRetryCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Retry", buttonreturnvalue:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbYesNoCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Yes", buttonreturnvalue:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="No", buttonreturnvalue:=vbNo
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbAbortRetryIgnore
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Abort", buttonreturnvalue:=vbAbort
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Retry", buttonreturnvalue:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton buttonrow:=lSetupRows, buttonindex:=lSetupRowButtons, buttoncaption:="Ignore", buttonreturnvalue:=vbIgnore
        Case Else
            MsgBox "The value provided for the ""buttons"" argument is not a known VB MsgBox value"
    End Select
    DsgnBttnsArea.Visible = True
    DsgnBttnRow(lSetupRows).Visible = True
    If Not AppliedBttnRows.Exists(DsgnBttnRow(lSetupRows)) Then AppliedBttnRows.Add DsgnBttnRow(lSetupRows), lSetupRows
    AppliedControls = DsgnBttnRow(lSetupRows)
    AppliedControls = DsgnBttnsFrame
    
xt: Exit Sub
eh: ErrMsg ErrSrc(PROC)
End Sub

Private Sub SetupMsgSect(ByVal msg_section As Long)
' -------------------------------------------------------------
' Setup a message section with its label when one is specified
' and return the message's width when greater than any other.
' Note: All height adjustments except the one for the text box
'       are done by the ResizeAndReposition
' -------------------------------------------------------------
    Const PROC = "SetupMsgSect"
    
    On Error GoTo eh
    Dim SectionMessage  As TypeMsgText:     SectionMessage = Me.MsgText(msg_section)
    Dim SectionLabel    As TypeMsgLabel:    SectionLabel = Me.MsgLabel(msg_section)
    Dim frArea          As MSForms.Frame:   Set frArea = DsgnMsgArea
    Dim frSection       As MSForms.Frame:   Set frSection = DsgnSection(msg_section)
    Dim la              As MSForms.Label:   Set la = DsgnSectionLabel(msg_section)
    Dim tbText          As MSForms.TextBox: Set tbText = DsgnSectionText(msg_section)
    Dim frText          As MSForms.Frame:   Set frText = DsgnTextFrame(msg_section)
        
    frSection.Width = frArea.Width
    la.Width = frSection.Width
    frText.Width = frSection.Width
    tbText.Width = frSection.Width
        
    If SectionMessage.Text <> vbNullString Then
    
        AppliedControls = frArea
        AppliedControls = frSection
        AppliedControls = frText
        AppliedControls = tbText
                
        If SectionLabel.Text <> vbNullString Then
            Set la = DsgnSectionLabel(msg_section)
            With la
                .Width = Me.InsideWidth - (siHmarginFrames * 2)
                .Caption = SectionLabel.Text
                With .Font
                    If SectionLabel.Monospaced Then
                        If SectionLabel.FontName <> vbNullString Then .name = SectionLabel.FontName Else .name = DEFAULT_LBL_MONOSPACED_FONT_NAME
                        If SectionLabel.FontSize <> 0 Then .Size = SectionLabel.FontSize Else .Size = DEFAULT_LBL_MONOSPACED_FONT_SIZE
                    Else
                        If SectionLabel.FontName <> vbNullString Then .name = SectionLabel.FontName Else .name = DEFAULT_LBL_PROPSPACED_FONT_NAME
                        If SectionLabel.FontSize <> 0 Then .Size = SectionLabel.FontSize Else .Size = DEFAULT_LBL_PROPSPACED_FONT_SIZE
                    End If
                    If SectionLabel.FontItalic Then .Italic = True
                    If SectionLabel.FontBold Then .Bold = True
                    If SectionLabel.FontUnderline Then .Underline = True
                End With
                If SectionLabel.FontColor <> 0 Then .ForeColor = SectionLabel.FontColor Else .ForeColor = rgbBlack
            End With
            frText.top = la.top + la.Height
            AppliedControls = la
        Else
            frText.top = 0
        End If
        
        If SectionMessage.Monospaced Then
            SetupMsgSectMonoSpaced msg_section  ' returns the maximum width required for monospaced section
        Else ' proportional spaced
            SetupMsgSectPropSpaced msg_section
        End If
        tbText.SelStart = 0
        
    End If
    
xt: Exit Sub
eh: ErrMsg ErrSrc(PROC)
End Sub

Private Sub SetupMsgSectMonoSpaced( _
                             ByVal msg_section As Long, _
                    Optional ByVal msg_append As Boolean = False, _
                    Optional ByVal msg_text As String = vbNullString)
' ------------------------------------------------------------------------------
' Setup the provided monospaced message section (msg_section). When a text is
' explicitely provided (msg_text) setup the sectiuon with this one, else with
' the property MsgText content. When an explicit text is provided the text
' either replaces the text, which the default or the text is appended when
' (msg_appen = True).
' Note 1: All top and height adjustments - except the one for the text box
'         itself are finally done by ResizeAndReposition services when all
'         elements had been set up.
' Note 2: The optional arguments (msg_append) and (msg_text) are used with the
'         Progress service which ma replace or add the provided text
' ------------------------------------------------------------------------------
Const PROC = "SetupMsgSectMonoSpaced"
    
    On Error GoTo eh
    Dim MsgSectText                 As TypeMsgText:     MsgSectText = MsgText(msg_section)
    Dim MsgAreaFrame                As MSForms.Frame:   Set MsgAreaFrame = DsgnMsgArea
    Dim MsgTextFrame                As MSForms.Frame:   Set MsgTextFrame = DsgnSectionTextFrame(msg_section)
    Dim MsgTextBox                  As MSForms.TextBox: Set MsgTextBox = DsgnSectionText(msg_section)
    Dim MsgSectFrame                As MSForms.Frame:   Set MsgSectFrame = DsgnSection(msg_section)
    Dim MsgAreaFrameMaxWidth        As Single
    Dim MsgSectFrameMaxWidth        As Single
    Dim MsgTextFrameMaxWidth        As Single
    Dim MsgSectFrameMaxWidthUsable  As Single
    
    '~~ Setup the textbox
    With MsgTextBox
        .Visible = True
        With .Font
            .Bold = MsgSectText.FontBold
            .Italic = MsgSectText.FontItalic
            .Underline = MsgSectText.FontUnderline
            If MsgSectText.FontSize <> 0 Then .Size = MsgSectText.FontSize Else .Size = DEFAULT_LBL_MONOSPACED_FONT_SIZE
            If MsgSectText.FontName <> vbNullString Then .name = MsgSectText.FontName Else .name = DEFAULT_LBL_MONOSPACED_FONT_NAME
        End With
        If MsgSectText.FontColor <> 0 Then .ForeColor = MsgSectText.FontColor Else .ForeColor = rgbBlack
        .MultiLine = True
        .WordWrap = False
        .AutoSize = True
        If msg_text = vbNullString Then
            .Value = MsgSectText.Text
        ElseIf msg_append Then
            If .Value = vbNullString Then .Value = msg_text Else .Value = .Value & vbLf & msg_text
        Else
            .Value = msg_text
        End If
        .AutoSize = False
        .SelStart = 0
        .Left = siHmarginFrames
        .Height = .Height + 2 ' ensure text is not squeeced
        MsgTextFrame.Left = siHmarginFrames
        MsgTextFrame.Width = .Width + (siHmarginFrames * 2)
        MsgSectFrame.Width = MsgTextFrame.Width + (siHmarginFrames * 2)
        MsgAreaFrame.Width = Max(MsgAreaFrame.Width, MsgSectFrame.Left + MsgSectFrame.Width + siHmarginFrames + SPACE_HORIZONTAL_SCROLLBAR)
        FormWidth = MsgAreaFrame.Width + siHmarginFrames + 7
        
        MsgAreaFrameMaxWidth = siMaxMsgWidthPts - (Me.Width - Me.InsideWidth) - siHmarginFrames
        MsgSectFrameMaxWidth = MsgAreaFrameMaxWidth - siHmarginFrames - SPACE_HORIZONTAL_SCROLLBAR
        MsgTextFrameMaxWidth = MsgSectFrameMaxWidth - siHmarginFrames
        MsgSectFrameMaxWidthUsable = siMaxMsgWidthPts - (siHmarginFrames * 2)
        '~~ The area width considers that there might be a need to apply a vertival scrollbar
        '~~ When the space finally isn't required, the sections are centered within the area
        FormWidth = MsgAreaFrame.Width + siHmarginFrames + 7
        
        If MsgTextFrame.Width > MsgTextFrameMaxWidth Then
            MsgSectFrame.Width = MsgSectFrameMaxWidth
            MsgAreaFrame.Width = MsgAreaFrameMaxWidth
            FormWidth = siMaxMsgWidthPts ' Reduce to specified maximum width
            '~~ Each monospaced message section will get its own horizontal scrollbar
            ApplyScrollBarHorizontal scroll_frame:=MsgTextFrame, new_width:=MsgTextFrameMaxWidth
        End If
        
    End With
    siMaxSectionWidth = Max(siMaxSectionWidth, MsgSectFrame.Width)
    
    '~~ Keep record of the controls which had been applied
    AppliedControls = MsgAreaFrame
    AppliedControls = MsgSectFrame
    AppliedControls = MsgTextFrame
    AppliedControls = MsgTextBox
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub SetupMsgSectPropSpaced( _
                             ByVal msg_section As Long, _
                    Optional ByVal msg_append As Boolean = False, _
                    Optional ByVal msg_text As String = vbNullString)
' ------------------------------------------------------------------------------
' Setup the provided section (msg_section) proportional spaced. When a text is
' explicitely provided (msg_text) setup the sectiuon with this one, else with
' the property MsgText content. When an explicit text is provided the text
' either replaces the text, which the default or the text is appended when
' (msg_appen = True).
' Note 1: When this proportional spaced section is setup the message width is
'         regarded final. However, top and height adjustments - except the one
'         for the text box itself are finally done by ResizeAndReposition
'         services when all elements had been set up.
' Note 2: The optional arguments (msg_append) and (msg_text) are used with the
'         Progress service which ma replace or add the provided text
' ------------------------------------------------------------------------------
    
    Dim MsgSectText As TypeMsgText:     MsgSectText = MsgText(msg_section)
    Dim frArea      As MSForms.Frame:   Set frArea = DsgnMsgArea
    Dim frSection   As MSForms.Frame:   Set frSection = DsgnSection(msg_section)
    Dim frText      As MSForms.Frame:   Set frText = DsgnSectionTextFrame(msg_section)
    Dim tbText      As MSForms.TextBox: Set tbText = DsgnSectionText(msg_section)
        
    '~~ For proportional spaced message sections the width is determined by the area width
    With frSection
        .Width = frArea.Width - siHmarginFrames - SPACE_HORIZONTAL_SCROLLBAR
        .Left = SPACE_HORIZONTAL_LEFT
        siMaxSectionWidth = Max(siMaxSectionWidth, .Width)
    End With
    With frText
        .Width = frSection.Width - siHmarginFrames
        .Left = SPACE_HORIZONTAL_LEFT
    End With
    
    With tbText
        .Visible = True
        With .Font
            If MsgSectText.FontName <> vbNullString Then .name = MsgSectText.FontName Else .name = DEFAULT_TXT_PROPSPACED_FONT_NAME
            If MsgSectText.FontSize <> 0 Then .Size = MsgSectText.FontSize Else .Size = DEFAULT_TXT_PROPSPACED_FONT_SIZE
            If MsgSectText.FontBold Then .Bold = True
            If MsgSectText.FontItalic Then .Italic = True
            If MsgSectText.FontUnderline Then .Underline = True
        End With
        If MsgSectText.FontColor <> 0 Then .ForeColor = MsgSectText.FontColor Else .ForeColor = rgbBlack
        DoEvents    ' This solves problems with invalid AutoSize adjustments !!!!!
        .MultiLine = True
        .AutoSize = True
        .WordWrap = True
        .Width = frText.Width - siHmarginFrames
        If msg_text = vbNullString Then
            .Value = MsgSectText.Text
        ElseIf msg_append Then
            If .Value = vbNullString Then .Value = msg_text Else .Value = .Value & vbLf & msg_text
        Else
            .Value = msg_text
        End If
        .AutoSize = True
'        .Width = frText.Width - siHmarginFrames
        .SelStart = 0
        .Left = SPACE_HORIZONTAL_LEFT
        frText.Width = .Left + .Width + siHmarginFrames
        DoEvents    ' to properly h-align the text
    End With
    
    AppliedControls = frArea
    AppliedControls = frSection
    AppliedControls = frText
    AppliedControls = tbText

End Sub

Private Sub SetupMsgSectsMonoSpaced()
    Const PROC = "SetupMsgSectsMonoSpaced"
    
    On Error GoTo eh
    Dim i       As Long
    Dim Message As TypeMsgText
    
    For i = 1 To Me.NoOfDesignedMsgSects
        Message = Me.MsgText(i)
        If Message.Monospaced And Message.Text <> vbNullString Then
            SetupMsgSect i
        End If
    Next i
    bDoneMonoSpacedSections = True
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub SetupMsgSectsPropSpaced()
    Const PROC = "SetupMsgSectsPropSpaced"
    
    On Error GoTo eh
    Dim i       As Long
    Dim Message As TypeMsgText
    
    For i = 1 To Me.NoOfDesignedMsgSects
        Message = MsgText(i)
        If Not Message.Monospaced And Message.Text <> vbNullString Then SetupMsgSect i
    Next i
    bDonePropSpacedSections = True
    bDoneMsgArea = True
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub SetupMsgSectUpdated(ByVal msg_section As Long)
' -----------------------------------------------------------
' Triggered by the Textbox_AfterUpdate event, adjusts the
' message form to the new sizeed Textbox.
' ------------------------------------------------------
'   Pending implementation
    msg_section = msg_section ' avoids not used message from MZTools
End Sub

Private Sub SetupTitle()
' ------------------------------------------------------------------------------------------
' When a specific font name and/or size is specified, the extra title label is actively used
' and the UserForm's title bar is not displayed - which means that there is no X to cancel.
' ------------------------------------------------------------------------------------------
    Const PROC = "SetupTitle"
    
    On Error GoTo eh
    Dim siTop           As Single
    
    siTop = 0
    With Me
        .Width = siMinMsgWidthPts ' Setup starts with the minimum message form width
        '~~ When a font name other then the standard UserForm font name is
        '~~ provided the extra hidden title label which mimics the title bar
        '~~ width is displayed. Otherwise it remains hidden.
        If sTitleFontName <> vbNullString And sTitleFontName <> .Font.name Then
            With .laMsgTitle   ' Hidden by default
                .Visible = True
                .top = siTop
                siTop = VgridPos(.top + .Height)
                .Font.name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .AutoSize = True
                .Caption = " " & sTitle    ' some left margin
                siTitleWidth = .Width + SPACE_HORIZONTAL_RIGHT
            End With
            AppliedControls = .laMsgTitle
            .laMsgTitleSpaceBottom.Visible = False
        Else
            '~~ The extra title label is only used to adjust the form width and remains hidden
            With .laMsgTitle
                With .Font
                    .Bold = False
                    .name = Me.Font.name
                    .Size = 8.65    ' Value which comes to a length close to the length required
                End With
                .Visible = True
                .Caption = vbNullString
                .AutoSize = True
                .Caption = " " & sTitle    ' some left margin
                siTitleWidth = .Width + 30
            End With
            .Caption = " " & sTitle    ' some left margin
            .laMsgTitleSpaceBottom.Visible = True
        End If
                
        .laMsgTitleSpaceBottom.Width = siTitleWidth
        FormWidth = siTitleWidth + (Me.Width - Me.InsideWidth)
    End With
    bDoneTitle = True
    
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub tbMsgSect1Text_AfterUpdate():    SetupMsgSectUpdated 1:   End Sub

Private Sub tbMsgSect2Text_AfterUpdate():    SetupMsgSectUpdated 2:   End Sub

Private Sub tbMsgSect3Text_AfterUpdate():    SetupMsgSectUpdated 3:   End Sub

Private Sub tbMsgSect4Text_AfterUpdate():    SetupMsgSectUpdated 4:   End Sub
'Private Sub tbMsgSect5Text_AfterUpdate():    SetupMsgSectUpdated 5:   End Sub

Private Sub UserForm_Activate()
' ---------------------------------------------------
' To avoid screen flicker the setup may has been done
' already.
' However for test purpose the Setup may run with the
' Activate event i.e. the .Show
' ---------------------------------------------------
    If bDoneSetup = True Then bDoneSetup = False Else Setup
End Sub

Public Function VgridPos(ByVal si As Single) As Single
' --------------------------------------------------------------
' Returns an integer of which the remainder (Int(si) / 6) is 0.
' Background: A controls content is only properly displayed
' when the top position of it is aligned to such a position.
' --------------------------------------------------------------
    Dim i As Long
    
    For i = 0 To 6
        If Int(si) = 0 Then
            VgridPos = 0
        Else
            If Int(si) < 6 Then si = 6
            If (Int(si) + i) Mod 6 = 0 Then
                VgridPos = Int(si) + i
                Exit For
            End If
        End If
    Next i

End Function

