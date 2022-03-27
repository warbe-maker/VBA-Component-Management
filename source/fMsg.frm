VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   11595
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   12390
   OleObjectBlob   =   "fMsg.frx":0000
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
' Public Properties:
' VisualizeControls
' IndicateFrameCaptions
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
' W. Rauschenberger Berlin, Jan 2022 (last revision)
' --------------------------------------------------------------------------
Const DFLT_BTTN_MIN_WIDTH           As Single = 70              ' Default minimum reply button width
Const DFLT_LBL_MONOSPACED_FONT_NAME As String = "Courier New"   ' Default monospaced font name
Const DFLT_LBL_MONOSPACED_FONT_SIZE As Single = 9               ' Default monospaced font size
Const DFLT_LBL_PROPSPACED_FONT_NAME As String = "Calibri"       ' Default proportional spaced font name
Const DFLT_LBL_PROPSPACED_FONT_SIZE As Single = 9               ' Default proportional spaced font size
Const DFLT_TXT_MONOSPACED_FONT_NAME As String = "Courier New"   ' Default monospaced font name
Const DFLT_TXT_MONOSPACED_FONT_SIZE As Single = 10              ' Default monospaced font size
Const DFLT_TXT_PROPSPACED_FONT_NAME As String = "Tahoma"        ' Default proportional spaced font name
Const DFLT_TXT_PROPSPACED_FONT_SIZE As Single = 10              ' Default proportional spaced font size
Const HSPACE_BTTN_AREA              As Single = 15              ' Minimum left and right margin for the centered buttons area
Const HSPACE_BUTTONS                As Single = 4               ' Horizontal space between reply buttons
Const HSPACE_LEFT                   As Single = 0               ' Left margin for labels and text boxes
Const HSPACE_RIGHT                  As Single = 15              ' Horizontal right space for labels and text boxes
Const HSPACE_LEFTRIGHT_BUTTONS      As Long = 8                 ' The margin before the left most and after the right most button
Const MARGIN_RIGHT_MSG_AREA         As String = 7
Const NEXT_ROW                      As String = vbLf            ' Reply button row break
Const SCROLL_V_WIDTH                As Single = 20              ' Additional horizontal space required for a frame with a vertical scrollbar
Const SCROLL_H_HEIGHT               As Single = 20              ' Additional vertical space required for a frame with a horizontal scroll barr
Const TEST_WITH_FRAME_BORDERS       As Boolean = False          ' For test purpose only! Display frames with visible border
Const TEST_WITH_FRAME_CAPTIONS      As Boolean = False          ' For test purpose only! Display frames with their test captions (erased by default)
Const VSPACE_AREAS                  As Single = 10              ' Vertical space between message area and replies area
Const VSPACE_BOTTOM                 As Single = 30              ' Space occupied by the title bar
Const VSPACE_BTTN_ROWS              As Single = 5               ' Vertical space between button rows
Const VSPACE_LABEL                  As Single = 0               ' Vertical space between the section-label and the following section-text
Const VSPACE_SECTIONS               As Single = 7               ' Vertical space between displayed message sections
Const VSPACE_TEXTBOXES              As Single = 18              ' Vertical bottom marging for all textboxes
Const VSPACE_TOP                    As Single = 2               ' Top position for the first displayed control
Const CTL_VISUALIZE_TBX_BACKCOL = &H80C0FF
Const CTL_VISUALIZE_TEXT_SECTION = &HFFFFC0

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
' -------------------------------------------------------------------------------

' For a much faster DoEvents alternative
'Private Declare PtrSafe Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
'Private Const QS_HOTKEY As Long = &H80
'Private Const QS_KEY As Long = &H1
'Private Const QS_MOUSEBUTTON As Long = &H4
'Private Const QS_PAINT As Long = &H20
' -------------------------------------------------------------------------------

' Timer means
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (TimerSystemFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
' -------------------------------------------------------------------------------

Private Enum enStartupPosition      ' ---------------------------
    sup_Manual = 0                  ' Used to position the
'    sup_CenterOwner = 1             ' final setup message form
'    sup_CenterScreen = 2            ' horizontally and vertically
'    sup_WindowsDefault = 3          ' centered on the screen
End Enum                            ' ---------------------------

Private Enum enMsgFormUsage
    usage_progress_display = 1
'    usage_message_display = 2
End Enum

Public Enum MSFormControls
    ' List of all the MSForms Controls.
    CheckBox
    ComboBox
    CommandButton
    Frame
    Image
    Label
    ListBox
    MultiPage
    OptionButton
    ScrollBar
    SpinButton
    TabStrip
    TextBox
    ToggleButton
End Enum

Private AppliedBttns            As Dictionary   ' Dictionary of applied buttons (key=CommandButton, item=row)
Private AppliedBttnsRetVal      As Dictionary   ' Dictionary of the applied buttons' reply value (key=CommandButton)
Private AreaBttns               As MSForms.Frame    ' Set with CollectDesignControls
Private AreaMsg                 As MSForms.Frame    ' Set with CollectDesignControls
Private bDoneMonoSpacedSects    As Boolean
Private bDoneMsgArea            As Boolean
Private bDonePropSpacedSects    As Boolean
Private bDoneTitle              As Boolean
Private bFormEvents             As Boolean
Private bIndicateFrameCaptions  As Boolean
Private bMonitorInitialized     As Boolean
Private bMonitorMode            As Boolean
Private bReplyWithIndex         As Boolean
Private bSetUpDone              As Boolean
Private BttnsFrm                As MSForms.Frame    ' Set with CollectDesignControls
Private bVisualizeControls      As Boolean
Private cllDsgnAreas            As Collection   ' Collection of the two primary/top frames
Private cllDsgnBttnRows         As Collection   ' Collection of the designed reply button row frames
Private cllDsgnBttns            As Collection   ' Collection of the collection of the designed reply buttons of a certain row
Private cllDsgnBttnsFrm         As Collection
Private cllDsgnMsgSects         As Collection   '
Private cllDsgnMsgSectsLabel    As Collection
Private cllDsgnMsgSectsTextBox  As Collection   ' Collection of section textboxes
Private cllDsgnMsgSectsTextFrm  As Collection   ' Collection of section textframes
Private cllDsgnRowBttns         As Collection   ' Collection of a designed reply button row's buttons
Private cllSteps                As Collection
Private cyTimerTicksBegin       As Currency
Private cyTimerTicksEnd         As Currency
Private dctMonoSpaced           As Dictionary
Private dctMonoSpacedTbx        As Dictionary
Private dctSectsLabel           As Dictionary   ' SectFrm specific label either provided via properties MsgLabel or Msg
Private dctSectsMonoSpaced      As Dictionary   ' SectFrm specific monospace option either provided via properties MsgMonospaced or Msg
Private dctSectsText            As Dictionary   ' SectFrm specific text either provided via properties MsgText or Msg
Private DfltMnSpcdLblFontName   As String
Private DfltMnSpcdLblFontSize   As Single
Private DfltMnSpcdTxtFontName   As String
Private DfltMnSpcdTxtFontSize   As Single
Private frmSteps                As MSForms.Frame
Private lBackColor              As Long
Private lMonitorSteps           As Long
Private lSetupRowButtons        As Long         ' number of buttons setup in a row
Private lSetupRows              As Long         ' number of setup button rows
Private lStepsDisplayed         As Long
Private MsgSectLbl              As TypeMsgText     ' Label section of the TypeMsg UDT
Private MsgSectTxt              As TypeMsgText      ' Text section of the TypeMsg UDT
Private MsgText1                As TypeMsgText  ' common text element
Private SectFrm                 As MSForms.Frame    ' Set with SetCtrlsOfSection for a certain section
Private Sections                As Long             ' Set with CollectDesignControls (number of message sections designed)
Private SectLbl                 As MSForms.Label    ' Set with SetCtrlsOfSection for a certain section
Private SectTxtBox              As MSForms.TextBox  ' Set with SetCtrlsOfSection for a certain section
Private SectTxtFrm              As MSForms.Frame    ' Set with SetCtrlsOfSection for a certain section
Private siHmarginButtons        As Single
Private siHmarginFrames         As Single       ' Test property, value defaults to 0
Private siMaxButtonHeight       As Single
Private siMaxButtonWidth        As Single
Private siMinButtonWidth        As Single
Private siMsgHeightMax          As Single       ' Maximum message height in pt
Private siMsgHeightMin          As Single       ' Minimum message height in pt
Private siMsgWidthMax           As Single       ' Maximum message width in pt
Private siMsgWidthMin           As Single       ' Minimum message width in pt
Private siVmarginButtons        As Single
Private siVmarginFrames         As Single       ' Test property, value defaults to 0
Private sMsgTitle               As String
Private tbxFooter               As MSForms.TextBox
Private tbxHeader               As MSForms.TextBox
Private tbxStep                 As MSForms.TextBox
Private TextLabel               As TypeMsgText
Private TextMonitorFooter       As TypeMsgText
Private TextMonitorHeader       As TypeMsgText
Private TextMonitorStep         As TypeMsgText
Private TextMsg                 As TypeMsgText
Private TextSection             As TypeMsg
Private TimerSystemFrequency    As Currency
Private TitleWidth              As Single
Private UsageType               As enMsgFormUsage
Private vButtons                As Variant
Private VirtualScreenHeightPts  As Single
Private VirtualScreenLeftPts    As Single
Private VirtualScreenTopPts     As Single
Private VirtualScreenWidthPts   As Single
Private vMsgButtonDefault       As Variant          ' Index or caption of the default button
Private vReplyValue             As Variant

Private Sub UserForm_Initialize()
    Const PROC = "UserForm_Initialize"
    
    On Error GoTo eh
    ' Get the display screen's dimensions and position in pts
    GetScreenMetrics VirtualScreenLeftPts _
                   , VirtualScreenTopPts _
                   , VirtualScreenWidthPts _
                   , VirtualScreenHeightPts
    
    If bSetUpDone Then GoTo xt
    Set dctMonoSpaced = New Dictionary
    Set dctMonoSpacedTbx = New Dictionary
    
    siMinButtonWidth = DFLT_BTTN_MIN_WIDTH
    siHmarginButtons = HSPACE_BUTTONS
    siVmarginButtons = VSPACE_BTTN_ROWS
    bFormEvents = False
    DfltMnSpcdTxtFontName = DFLT_TXT_MONOSPACED_FONT_NAME
    DfltMnSpcdTxtFontSize = DFLT_TXT_MONOSPACED_FONT_SIZE
    DfltMnSpcdLblFontName = DFLT_LBL_MONOSPACED_FONT_NAME
    DfltMnSpcdLblFontSize = DFLT_LBL_MONOSPACED_FONT_SIZE
    siHmarginFrames = 0     ' Ensures proper command buttons framing, may be used for test purpose
    Me.VmarginFrames = 0    ' Ensures proper command buttons framing and vertical positioning of controls
    SetupDone = False
    bDoneTitle = False
    bDoneMonoSpacedSects = False
    bDonePropSpacedSects = False
    bDoneMsgArea = False
    vMsgButtonDefault = 1
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub UserForm_Terminate()
    Set AppliedBttns = Nothing
    Set AppliedBttnsRetVal = Nothing
    Set cllDsgnAreas = Nothing
    Set cllDsgnBttnRows = Nothing
    Set cllDsgnBttns = Nothing
    Set cllDsgnBttnsFrm = Nothing
    Set cllDsgnMsgSects = Nothing
    Set cllDsgnMsgSectsLabel = Nothing
    Set cllDsgnMsgSectsTextBox = Nothing
    Set cllDsgnMsgSectsTextFrm = Nothing
    Set cllDsgnRowBttns = Nothing
    Set dctMonoSpaced = Nothing
    Set dctMonoSpacedTbx = Nothing
    Set dctSectsLabel = Nothing
    Set dctSectsMonoSpaced = Nothing
    Set dctSectsText = Nothing
End Sub

Private Property Get AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton) As Variant
    AppliedButtonRetVal = AppliedBttnsRetVal(Button)
End Property

Private Property Let AppliedButtonRetVal(Optional ByVal Button As MSForms.CommandButton, ByVal v As Variant)
    AppliedBttnsRetVal.Add Button, v
End Property

Private Property Get AppliedButtonRowHeight() As Single
    AppliedButtonRowHeight = siMaxButtonHeight + 2
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

Private Property Get DsgnBttn(Optional ByVal bttn_row As Long, Optional ByVal bttn_no As Long) As MSForms.CommandButton
    Set DsgnBttn = cllDsgnBttns(bttn_row)(bttn_no)
End Property

Private Property Get DsgnBttnRow(Optional ByVal lRow As Long) As MSForms.Frame:              Set DsgnBttnRow = cllDsgnBttnRows(lRow):                             End Property

Private Property Get DsgnBttnRows() As Collection:                                          Set DsgnBttnRows = cllDsgnBttnRows:                                 End Property

Private Property Get DsgnBttnsFrm() As MSForms.Frame:                                     Set DsgnBttnsFrm = cllDsgnBttnsFrm(1):                          End Property

Private Property Get DsgnMsgSect(Optional msg_section As Long) As MSForms.Frame:            Set DsgnMsgSect = cllDsgnMsgSects(msg_section):                     End Property

Private Property Get DsgnMsgSectLbl(Optional msg_section As Long) As MSForms.Label:       Set DsgnMsgSectLbl = cllDsgnMsgSectsLabel(msg_section):           End Property

Private Property Get DsgnMsgSects() As Collection:                                          Set DsgnMsgSects = cllDsgnMsgSects:                                 End Property

Private Property Get DsgnMsgSectsTextFrame() As Collection:                                 Set DsgnMsgSectsTextFrame = cllDsgnMsgSectsTextFrm:                End Property

Private Property Get DsgnMsgSectTxtBox(Optional msg_section As Long) As MSForms.TextBox:   Set DsgnMsgSectTxtBox = cllDsgnMsgSectsTextBox(msg_section):       End Property

Private Property Get DsgnMsgSectTxtFrm(Optional ByVal msg_section As Long):              Set DsgnMsgSectTxtFrm = cllDsgnMsgSectsTextFrm(msg_section):    End Property

Private Property Let FormWidth(ByVal considered_width As Single)
' ------------------------------------------------------------------------------
' The FormWidth property ensures
' - it is not less than the minimum specified width
' - it does not exceed the specified or the default maximum value
' - it may expand up to the maximum but never shrink
' ------------------------------------------------------------------------------
    Dim new_width As Single
    new_width = Max(Me.Width, TitleWidth, siMsgWidthMin, considered_width + 15)
    Me.Width = Min(new_width, siMsgWidthMax + Max(ScrollV_Width(AreaMsg), ScrollV_Width(AreaBttns)))
End Property

Private Property Get FormWidthMaxUsable()
    FormWidthMaxUsable = siMsgWidthMax - 15
End Property

Public Property Let HmarginButtons(ByVal si As Single):                                     siHmarginButtons = si:                                              End Property

Public Property Let HmarginFrames(ByVal si As Single):                                      siHmarginFrames = si:                                               End Property

Public Property Let IndicateFrameCaptions(ByVal b As Boolean):                              bIndicateFrameCaptions = b:                                         End Property

Public Property Get IsApplied(Optional ByVal v As Variant) As Boolean
    Dim frm As MSForms.Frame
    Dim tbx As MSForms.TextBox
    Dim cmb As MSForms.CommandButton
    Dim lbl As MSForms.Label
    
    Select Case TypeName(v)
        Case "Frame":           Set frm = v:    IsApplied = frm.Visible
        Case "TextBox":         Set tbx = v:    IsApplied = tbx.Visible
        Case "CommandButton":   Set cmb = v:    IsApplied = cmb.Visible
        Case "Label":           Set lbl = v:    IsApplied = lbl.Visible
    End Select
End Property

Public Property Let IsApplied(Optional ByVal v As Variant, ByVal b As Boolean)
    Const PROC = "IsApplied_Let"
    
    On Error GoTo eh
    Dim frm As MSForms.Frame
    Dim tbx As MSForms.TextBox
    Dim cmb As MSForms.CommandButton
    Dim lbl As MSForms.Label
    
    Select Case TypeName(v)
        Case "Frame":           Set frm = v:    frm.Visible = b
        Case "TextBox":         Set tbx = v:    tbx.Visible = b
        Case "CommandButton":   Set cmb = v:    cmb.Visible = b
        Case "Label":           Set lbl = v:    lbl.Visible = b
    End Select

xt: Exit Property

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Private Property Get MaxRowsHeight() As Single:                                             MaxRowsHeight = siMaxButtonHeight + (siVmarginFrames * 2):          End Property

'Private Property Get MaxWidthAreaBttns() As Single
'    MaxWidthAreaBttns = FormWidthMaxUsable - (HSPACE_BTTN_AREA * 2)
'End Property

Private Property Get MaxWidthMsgArea() As Single
' ------------------------------------------------------------------------------
' The maximum usable message area width considers the specified maximum form
' width and the InsideWidth
' ------------------------------------------------------------------------------
    MaxWidthMsgArea = Me.InsideWidth
End Property

Public Property Let MinButtonWidth(ByVal si As Single):                                     siMinButtonWidth = si:                                                          End Property

Public Property Let MonitorMode(ByVal b As Boolean):                                        bMonitorMode = b:                                              End Property

Private Property Get MonoSpaced(Optional ByVal var_ctl As Variant) As Boolean
    MonoSpaced = dctMonoSpaced.Exists(var_ctl)
End Property

Private Property Let MonoSpaced( _
                 Optional ByVal var_ctl As Variant, _
                          ByVal b As Boolean)
    If b Then
        If Not dctMonoSpaced.Exists(var_ctl) Then dctMonoSpaced.Add var_ctl, var_ctl.name
    Else
        If dctMonoSpaced.Exists(var_ctl) Then dctMonoSpaced.Remove var_ctl
    End If
End Property

Private Property Get MonoSpacedTbx(Optional ByVal tbx As MSForms.TextBox) As Boolean
    MonoSpacedTbx = dctMonoSpacedTbx.Exists(tbx)
End Property

Private Property Let MonoSpacedTbx( _
                 Optional ByVal tbx As MSForms.TextBox, _
                          ByVal b As Boolean)
    If b Then
        If Not dctMonoSpacedTbx.Exists(tbx) Then dctMonoSpacedTbx.Add tbx, tbx.name
    Else
        If dctMonoSpacedTbx.Exists(tbx) Then dctMonoSpacedTbx.Remove tbx
    End If
End Property

Private Property Get MSFormsProgID(Optional mfc As MSFormControls) As String
' ------------------------------------------------------------------------------
' Returns the ProgID for the control (mfc). See
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms
' ------------------------------------------------------------------------------
    Select Case mfc
      Case MSFormControls.CheckBox:       MSFormsProgID = "Forms.CheckBox.1"
      Case MSFormControls.ComboBox:       MSFormsProgID = "Forms.ComboBox.1"
      Case MSFormControls.CommandButton:  MSFormsProgID = "Forms.CommandButton.1"
      Case MSFormControls.Frame:          MSFormsProgID = "Forms.Frame.1"
      Case MSFormControls.Image:          MSFormsProgID = "Forms.Image.1"
      Case MSFormControls.Label:          MSFormsProgID = "Forms.Label.1"
      Case MSFormControls.ListBox:        MSFormsProgID = "Forms.ListBox.1"
      Case MSFormControls.MultiPage:      MSFormsProgID = "Forms.MultiPage.1"
      Case MSFormControls.OptionButton:   MSFormsProgID = "Forms.OptionButton.1"
      Case MSFormControls.ScrollBar:      MSFormsProgID = "Forms.ScrollBar.1"
      Case MSFormControls.SpinButton:     MSFormsProgID = "Forms.SpinButton.1"
      Case MSFormControls.TabStrip:       MSFormsProgID = "Forms.TabStrip.1"
      Case MSFormControls.TextBox:        MSFormsProgID = "Forms.TextBox.1"
      Case MSFormControls.ToggleButton:   MSFormsProgID = "Forms.ToggleButton.1"
    End Select
End Property

Public Property Let MsgButtonDefault(ByVal vDefault As Variant)
    vMsgButtonDefault = vDefault
End Property

Public Property Let MsgButtons(ByVal v As Variant)
        
    Select Case VarType(v)
        Case vbLong, vbString:  vButtons = v
        Case vbEmpty:           vButtons = vbOKOnly
        Case Else
            If IsArray(v) Then
                vButtons = v
            ElseIf TypeName(v) = "Collection" Or TypeName(v) = "Dictionary" Then
                Set vButtons = v
            End If
    End Select
End Property

Public Property Get MsgHeightMax() As Single:           MsgHeightMax = siMsgHeightMax:  End Property

Public Property Let MsgHeightMax(ByVal si As Single):   siMsgHeightMax = si:            End Property

Public Property Get MsgHeightMin() As Single:           MsgHeightMin = siMsgHeightMin:  End Property

Public Property Let MsgHeightMin(ByVal si As Single):   siMsgHeightMin = si:            End Property

Public Property Get MsgLabel(Optional ByVal msg_section As Long) As TypeMsgText
' ------------------------------------------------------------------------------
' Transfers a section's message UDT stored as array back into the UDT.
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    If dctSectsLabel Is Nothing Then
        MsgLabel.Text = vbNullString
    Else
        If dctSectsLabel.Exists(msg_section) Then
            vArry = dctSectsLabel(msg_section)
            MsgLabel.FontBold = vArry(0)
            MsgLabel.FontColor = vArry(1)
            MsgLabel.FontItalic = vArry(2)
            MsgLabel.FontName = vArry(3)
            MsgLabel.FontSize = vArry(4)
            MsgLabel.FontUnderline = vArry(5)
            MsgLabel.MonoSpaced = vArry(6)
            MsgLabel.Text = vArry(7)
        Else
            MsgLabel.Text = vbNullString
        End If
    End If
End Property

Public Property Let MsgLabel(Optional ByVal msg_section As Long, _
                                      ByRef msg_label As TypeMsgText)
' ------------------------------------------------------------------------------
' Transfers a message label UDT (msg_label) into an array and stores it in the
' Dictionary (dctSectsLabel) with the section (msg_section) as the key.
' ------------------------------------------------------------------------------
    Dim vArry(7)    As Variant
    
    If dctSectsLabel Is Nothing Then Set dctSectsLabel = New Dictionary
    If Not dctSectsLabel.Exists(msg_section) Then
        vArry(0) = msg_label.FontBold
        vArry(1) = msg_label.FontColor
        vArry(2) = msg_label.FontItalic
        vArry(3) = msg_label.FontName
        vArry(4) = msg_label.FontSize
        vArry(5) = msg_label.FontUnderline
        vArry(6) = msg_label.MonoSpaced
        vArry(7) = msg_label.Text
        dctSectsLabel.Add msg_section, vArry
    End If
End Property

Public Property Get MsgMonoSpaced(Optional ByVal msg_section As Long) As Boolean
    Dim vArry() As Variant
    
    If dctSectsText Is Nothing Then
        MsgMonoSpaced = False
    Else
        With dctSectsText
            If .Exists(msg_section) Then
                vArry = .Item(msg_section)
                MsgMonoSpaced = vArry(6)
            Else
                MsgMonoSpaced = False
            End If
        End With
    End If
End Property

Public Property Get MsgText(Optional ByVal msg_section As Long) As TypeMsgText
' ------------------------------------------------------------------------------
' Transferes message UDT stored as array back into the UDT.
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    
    If dctSectsText Is Nothing Then
        MsgText.Text = vbNullString
    Else
        If dctSectsText.Exists(msg_section) Then
            vArry = dctSectsText(msg_section)
            MsgText.FontBold = vArry(0)
            MsgText.FontColor = vArry(1)
            MsgText.FontItalic = vArry(2)
            MsgText.FontName = vArry(3)
            MsgText.FontSize = vArry(4)
            MsgText.FontUnderline = vArry(5)
            MsgText.MonoSpaced = vArry(6)
            MsgText.Text = vArry(7)
        Else
            MsgText.Text = vbNullString
        End If
    End If

End Property

Public Property Let MsgText(Optional ByVal msg_section As Long, _
                                     ByRef msg_msg As TypeMsgText)
' ------------------------------------------------------------------------------
' Transfers a message UDT into an array and stores it in a Dictionary with the
' section number (msg_section) as the key.
' ------------------------------------------------------------------------------
    Dim vArry(7)    As Variant
    
    If dctSectsText Is Nothing Then Set dctSectsText = New Dictionary
    If Not dctSectsText.Exists(msg_section) Then
        vArry(0) = msg_msg.FontBold
        vArry(1) = msg_msg.FontColor
        vArry(2) = msg_msg.FontItalic
        vArry(3) = msg_msg.FontName
        vArry(4) = msg_msg.FontSize
        vArry(5) = msg_msg.FontUnderline
        vArry(6) = msg_msg.MonoSpaced
        vArry(7) = msg_msg.Text
        dctSectsText.Add msg_section, vArry
    End If
End Property

Public Property Get MsgTitle() As String:               MsgTitle = Me.Caption:                                          End Property

Public Property Let MsgTitle(ByVal s As String):        sMsgTitle = s:                                                  End Property

Public Property Get MsgWidthMax() As Single:            MsgWidthMax = siMsgWidthMax:                                    End Property

Public Property Let MsgWidthMax(ByVal si As Single):    siMsgWidthMax = si:                                             End Property

Public Property Get MsgWidthMin() As Single:            MsgWidthMin = siMsgWidthMin:                                    End Property

Public Property Let MsgWidthMin(ByVal si As Single):    siMsgWidthMin = si:                                             End Property

Public Property Let NewHeight(Optional ByRef nh_frame_or_form As Object, _
                              Optional ByVal nh_for_visible_only As Boolean = True, _
                              Optional ByVal nh_y_action As fmScrollAction = fmScrollActionBegin, _
                                       ByVal nh_height As Single)
' ------------------------------------------------------------------------------
' Mimics a height change event. When the height of the frame (nh_frame_or_form) is
' changed (nh_frame_or_form_height) to less than the frame's content height and no vertical
' scrollbar is applied one is applied with the frame content's height. If one
' is already applied just the height is adjusted to the frame content's height.
' When the height becomes more than the frame's content height a vertical
' scrollbar becomes obsolete and is removed.
' ------------------------------------------------------------------------------
    Const PROC          As String = "NewHeight"
    
    On Error GoTo eh
    Dim siHeight    As Single
    
    If nh_frame_or_form Is Nothing Then Err.Raise AppErr(1), ErrSrc(PROC), "The required argument 'nh_frame_or_form' is Nothing!"
    If Not IsFrameOrForm(nh_frame_or_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
    
    siHeight = ContentHeight(nh_frame_or_form, nh_for_visible_only)  ' height of the content in the frame or UserForm
    nh_frame_or_form.Height = nh_height
    
    If nh_frame_or_form.Height < siHeight Then
        '~~ Apply a vertical scrollbar if none is applied yet, adjust its height otherwise
        If Not ScrollV_Applied(nh_frame_or_form) Then
            ScrollV_Apply sa_height:=siHeight, sa_frame_or_form:=nh_frame_or_form, sa_y_action:=nh_y_action
        Else
            nh_frame_or_form.ScrollHeight = siHeight
            nh_frame_or_form.Scroll yAction:=nh_y_action
        End If
    End If
'    nh_frame_or_form.Height = ContentHeight(nh_frame_or_form, nh_for_visible_only)
    
xt: Exit Property
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Public Property Let NewWidth(Optional ByRef nw_frame_or_form As Object, _
                             Optional ByVal nw_for_visible_only As Boolean = True, _
                                      ByVal nw_width As Single)
' ------------------------------------------------------------------------------
' Applies a horizontal scroll-bar when the new width (nw_width) is less than
' the content's width. When a horizontal scroll-bar is already applied only its
' is applied (when none is applied yet) or the already applied one's width is
' width is adjusted to the content's width.
' When no frame (nw_frame_or_form) is provided, it is the Userform which is concerned.
' ------------------------------------------------------------------------------
    Const PROC = "NewWidth"
    
    On Error GoTo eh
    Dim siContentWidth  As Single
    
    If nw_frame_or_form Is Nothing Then Err.Raise AppErr(1), ErrSrc(PROC), "The required argument 'nw_frame_or_form' is Nothing!"
    If Not IsFrameOrForm(nw_frame_or_form) Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided argument 'nw_frame_or_form' is neither a Frame nor a Form!"
    
    siContentWidth = ContentWidth(nw_frame_or_form, nw_for_visible_only) - ScrollV_Width(nw_frame_or_form)
    
    With nw_frame_or_form
        .Width = nw_width
        If nw_width < siContentWidth Then
            '~~ Apply a horizontal scrollbar if none is applied yet,
            '~~ or just adjust its width otherwise
            If Not ScrollH_Applied(nw_frame_or_form) Then
                ScrollH_Apply sa_width:=siContentWidth, sa_frame_or_form:=nw_frame_or_form
            Else
                .ScrollWidth = siContentWidth
            End If
        End If
'        .Width = ContentWidth(nw_frame_or_form, nw_for_visible_only)
    End With
    
xt: Exit Property
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Public Property Get NoOfDesignedMsgSects() As Long ' -----------------------
    NoOfDesignedMsgSects = 4                       ' Global definition !!!!!
End Property                                       ' -----------------------

Private Property Get PrcntgHeightAreaBttns() As Single
    PrcntgHeightAreaBttns = Round(AreaBttns.Height / (AreaMsg.Height + AreaBttns.Height), 2)
End Property

Private Property Get PrcntgHeightMsgArea() As Single
    PrcntgHeightMsgArea = Round(AreaMsg.Height / (AreaMsg.Height + AreaBttns.Height), 2)
End Property

Public Property Get ReplyValue() As Variant:                ReplyValue = vReplyValue:                                   End Property

Public Property Let ReplyWithIndex(ByVal b As Boolean):     bReplyWithIndex = b:                                        End Property

Public Property Let SetupDone(ByVal b As Boolean):          bSetUpDone = b:         End Property

Private Property Get SysFrequency() As Currency
    If TimerSystemFrequency = 0 Then getFrequency TimerSystemFrequency
    SysFrequency = TimerSystemFrequency
End Property

Public Property Get Text(Optional ByVal txt_kind_of_text As KindOfText, _
                         Optional ByVal txt_section As Long = 1) As TypeMsgText
' ------------------------------------------------------------------------------
' Returns the provided text as section-text or -label, monitor-header,
' -footer, or -step.
' ------------------------------------------------------------------------------
    Select Case txt_kind_of_text
        Case m_header:    Text = TextMonitorHeader
        Case m_footer:    Text = TextMonitorFooter
        Case m_step:      Text = TextMonitorStep
        Case m_text:      TextSection.Section(txt_section).Text = TextMsg
        Case m_label:     TextSection.Section(txt_section).Label = TextLabel
    End Select
    
End Property

Public Property Let Text(Optional ByVal txt_kind_of_text As KindOfText, _
                         Optional ByVal txt_section As Long = 1, _
                                  ByRef txt_text As TypeMsgText)
' ------------------------------------------------------------------------------
' Provide the text (txt_text) as section (txt_section) text, section label,
' monitor header, footer, or step (txt_kind_of_text).
' ------------------------------------------------------------------------------
    Dim t As TypeMsgText
    t.FontBold = txt_text.FontBold
    t.FontColor = txt_text.FontColor
    t.FontItalic = txt_text.FontItalic
    t.FontName = txt_text.FontName
    t.FontSize = txt_text.FontSize
    t.FontUnderline = txt_text.FontUnderline
    t.MonoSpaced = txt_text.MonoSpaced
    t.Text = txt_text.Text
    Select Case txt_kind_of_text
        Case m_header:    TextMonitorHeader = t
        Case m_footer:    TextMonitorFooter = t
        Case m_step:      TextMonitorStep = t
        Case m_text:      TextSection.Section(txt_section).Text = t
        Case m_label:     TextSection.Section(txt_section).Label = t
    End Select

End Property

Private Property Get TimerSecsElapsed() As Currency:        TimerSecsElapsed = TimerTicksElapsed / SysFrequency:        End Property

Private Property Get TimerSysCurrentTicks() As Currency:    getTickCount TimerSysCurrentTicks:  End Property

Private Property Get TimerTicksElapsed() As Currency:       TimerTicksElapsed = cyTimerTicksEnd - cyTimerTicksBegin:    End Property

Public Property Get VisualizeControls() As Boolean:                                         VisualizeControls = bVisualizeControls:                             End Property

Public Property Let VisualizeControls(ByVal b As Boolean):                                  bVisualizeControls = b:                                             End Property

Public Property Get VmarginButtons() As Single:             VmarginButtons = siVmarginButtons:                          End Property

Public Property Let VmarginButtons(ByVal si As Single):     siVmarginButtons = si:                                      End Property

Public Property Get VmarginFrames() As Single:              VmarginFrames = siVmarginFrames:                            End Property

Public Property Let VmarginFrames(ByVal si As Single):      siVmarginFrames = AdjustToVgrid(si):                             End Property

Public Function AddControl(ByVal ac_ctl As MSFormControls _
                , Optional ByVal ac_in As MSForms.Frame = Nothing _
                , Optional ByVal ac_name As String = vbNullString _
                , Optional ByVal ac_visible As Boolean = False _
                          ) As MSForms.Control
' ------------------------------------------------------------------------------
' Returns the control (ac_ctl) added to the to the userform or a frame (ac_in),
' optionally named (ac_name) and by default invisible (ac_visible).
' ------------------------------------------------------------------------------
    If ac_in Is Nothing _
    Then Set AddControl = Me.Controls.Add(MSFormsProgID(ac_ctl), ac_name, ac_visible) _
    Else Set AddControl = ac_in.Controls.Add(MSFormsProgID(ac_ctl), ac_name, ac_visible)
End Function

Private Sub AdjustedParentsWidthAndHeight(ByVal ctrl As MSForms.Control)
' ------------------------------------------------------------------------------
' Adjust the width and height of the parent frame of the control therein (ctrl)
' by considering the control's width and height and a possibly applied vertical
' and/or horizontal scroll-bar to the parent frame.
' ------------------------------------------------------------------------------
    Dim FrmParent   As Variant
    Dim VScroll     As Boolean
    Dim HScroll     As Boolean
    Dim AddWidth    As Single
    Dim AddHeight   As Single
    
    AddWidth = SCROLL_V_WIDTH
    AddHeight = SCROLL_H_HEIGHT
    On Error Resume Next
    Set FrmParent = ctrl.Parent
    If Err.Number <> 0 Then
        On Error GoTo eh
        GoTo xt
    End If
    
    While Err.Number = 0
        VScroll = ScrollV_Applied(FrmParent)
        HScroll = ScrollH_Applied(FrmParent)
'        Debug.Print "V" & Abs(CInt(VScroll)) & " H" & Abs(CInt(HScroll)) & " " & FrmParent.Name
'        If VScroll Then Stop
        With ctrl
            If VScroll And Not HScroll Then
                FrmParent.Width = .Left + .Width + AddWidth
                AddWidth = 0
            ElseIf HScroll And Not VScroll Then
                FrmParent.Height = .Top + .Height + AddHeight
                AddHeight = 0
            ElseIf Not VScroll And Not HScroll Then
                Select Case TypeName(ctrl)
                    Case "TextBox"
                        FrmParent.Width = .Left + .Width + 2
                        FrmParent.Height = .Top + .Height
                    Case "Frame"
                        If ctrl Is AreaMsg Then
                            FrmParent.Width = .Left + .Width + 2
                            FrmParent.Height = .Top + ContentHeight(ctrl)
                        Else
                            FrmParent.Width = .Left + .Width + 2
                            FrmParent.Height = .Top + .Height
                        End If
                    Case Else
                        FrmParent.Width = .Left + .Width + 10
                        FrmParent.Height = .Top + ContentHeight(ctrl)
                End Select
            End If
        End With
        Set ctrl = FrmParent
        Set FrmParent = FrmParent.Parent
    Wend
                
xt: '~~ Adjust finally the top frame's width and height which is the UserForm
    FrmParent.Height = ctrl.Top + ContentHeight(FrmParent) + 20
    FrmParent.Width = ContentWidth(FrmParent) + 18
    Exit Sub
eh:
End Sub

Private Sub AdjustTopPositions()
' ------------------------------------------------------------------------------
' - Adjusts each visible control's top position considering its current height.
' ------------------------------------------------------------------------------
    Const PROC = "AdjustTopPositions"
    
    On Error GoTo eh
    Dim i                   As Long
    Dim MaxTextFrameWidth   As Single
    Dim TopPosTextFrame     As Single
    Dim TopPosNextSect      As Single
    
    MaxTextFrameWidth = MaxUsedWidthTextFrames
    TopPosNextSect = 0 ' The top sections top position
    AreaMsg.Top = 0
    
    For i = 1 To Sections
        TopPosTextFrame = 0
        SetCtrlsOfSection i
        If IsApplied(SectFrm) Then
            '~~ Adjust top positions of message section and its items
            
            '~~ Top pos section label
            If IsApplied(SectLbl) Then
                '~~ Adjust the section label
                With SectLbl
                    .Top = 0
                    TopPosTextFrame = AdjustToVgrid(.Top + .Height)
                    SectLbl.Width = Me.Width - .Left - 5
                End With
            End If
            
            '~~ Top pos TextBox
            SectTxtBox.Top = 0
            '~~ Top pos Text Frasme
            SectTxtFrm.Top = TopPosTextFrame
            
            '~~ Top pos Message Section
            With SectFrm
                .Top = TopPosNextSect
                .Width = MaxTextFrameWidth
                If IsApplied(SectTxtFrm) Then
                    .Height = SectTxtFrm.Top + SectTxtFrm.Height
                End If
                TopPosNextSect = AdjustToVgrid(.Top + .Height + VSPACE_SECTIONS)
            End With
            TimedDoEvents    ' to properly h-align the text
            AdjustedParentsWidthAndHeight SectFrm
        End If
    Next i
    
    '~~ Top position Message Area
    If IsApplied(AreaBttns) And IsApplied(AreaMsg) Then
        AreaBttns.Top = AreaMsg.Top + AreaMsg.Height + VSPACE_AREAS
        Me.Height = AreaBttns.Top + AreaBttns.Height + VSPACE_AREAS
    
    ElseIf IsApplied(AreaBttns) And Not IsApplied(AreaMsg) Then
        AreaBttns.Top = VSPACE_AREAS
        FrameCenterHorizontal AreaBttns
        Me.Height = AreaBttns.Top + AreaBttns.Height + VSPACE_AREAS
    
    ElseIf Not IsApplied(AreaBttns) And IsApplied(AreaMsg) Then
        Me.Height = AreaMsg.Top + AreaMsg.Height + VSPACE_AREAS
    End If
    Me.Height = Me.Height + VSPACE_BOTTOM

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MonitorStepsAdjustTopPosition()
' ------------------------------------------------------------------------------
' - Adjusts each visible control's top position considering its current height.
' ------------------------------------------------------------------------------
    Const PROC = "AdjustTopPositions"
    
    On Error GoTo eh
    Dim siTop   As Single
    Dim ctl     As MSForms.Control
    Dim v       As Variant
    
    siTop = 0
    For Each v In cllSteps
        Set ctl = v
        ctl.Top = siTop
        siTop = AdjustToVgrid(ctl.Top + ctl.Height)
    Next v

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub





Public Function AdjustToVgrid(ByVal atvg_si As Single, _
                     Optional ByVal atvg_threshold As Single = 1.5, _
                     Optional ByVal atvg_grid As Single = 6) As Single
' -------------------------------------------------------------------------------
' Returns an integer which is a multiple of the grid value (stvg_grid) - which
' defaults to 6 - by considering a certain threshold (atvg_threshold) - which
' defaults to 1.5.
' The function is used to vertically align form controls with the grid in order
' result vertically aligns a control in a userform to a grid value which ensures
' to have any text within the control correctly displayed in accordance with its
' font size. A certain threshold prevents an optically irritating large space to
' a control abovel. Examples:
'  7.5 < si >= 0   results to 6
' 13.5 < si >= 7.5 results in 12
' When the provided atvg_si is not a type single it is asumed it is a control and
' the value is calculated .top + .height.
' -------------------------------------------------------------------------------
    AdjustToVgrid = (Int((atvg_si - atvg_threshold) / atvg_grid) * atvg_grid) + atvg_grid
End Function

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

Private Function AppliedBttnRows() As Dictionary
' ------------------------------------------------------------------------------
' Returns a Dictionary of the applied/used/visible butoon rows with the row
' frame as the key and the applied/visible buttons therein as item.
' ------------------------------------------------------------------------------
    
    Dim dct             As New Dictionary
    Dim ButtonRows      As Long
    Dim ButtonRowsRow   As MSForms.Frame
    Dim v               As Variant
    Dim ButtonsInRow    As Long
    
    For ButtonRows = 1 To DsgnBttnRows.Count
        Set ButtonRowsRow = DsgnBttnRows(ButtonRows)
        If IsApplied(ButtonRowsRow) Then
            ButtonsInRow = 0
            For Each v In DsgnRowBttns(ButtonRows)
                If IsApplied(v) Then ButtonsInRow = ButtonsInRow + 1
            Next v
            dct.Add ButtonRowsRow, ButtonsInRow
        End If
    Next ButtonRows
    Set AppliedBttnRows = dct

End Function

Public Sub AutoSizeTextBox(ByRef as_tbx As MSForms.TextBox, _
                           ByVal as_text As String, _
                  Optional ByVal as_width_limit As Single = 0, _
                  Optional ByVal as_width_min As Single = 0, _
                  Optional ByVal as_height_min As Single = 0, _
                  Optional ByVal as_width_max As Single = 0, _
                  Optional ByVal as_height_max As Single = 0, _
                  Optional ByVal as_append As Boolean = False, _
                  Optional ByVal as_append_margin As String = vbNullString)
' ------------------------------------------------------------------------------
' Common AutoSize service for an MsForms.TextBox providing a width and height
' for the TextBox (as_tbx) by considering:
' - When a width limit is provided (as_width_limit > 0) the width is regarded a
'   fixed maximum and thus the height is auto-sized by means of WordWrap=True.
' - When no width limit is provided (the default) WordWrap=False and thus the
'   width of the TextBox is determined by the longest line.
' - When a maximum width is provided (as_width_max > 0) and the parent of the
'   TextBox is a frame a horizontal scrollbar is applied for the parent frame.
' - When a maximum height is provided (as_heightmax > 0) and the parent of the
'   TextBox is a frame a vertical scrollbar is applied for the parent frame.
' - When a minimum width (as_width_min > 0) or a minimum height (as_height_min
'   > 0) is provided the size of the textbox is set correspondingly. This
'   option is specifically usefull when text is appended to avoid much flicker.
'
' Uses: NewWidth, FrameContentWidth, ScrollH_Apply,
'       NewHeight, ContentHeight, ScrollV_Apply
'
' W. Rauschenberger Berlin June 2021
' ------------------------------------------------------------------------------
    
    With as_tbx
        .MultiLine = True
        If as_width_limit > 0 Then
            '~~ AutoSize the height of the TextBox considering the limited width
            .WordWrap = True
            .AutoSize = False
            .Width = as_width_limit - 2 ' the readability space is added later
            
            If Not as_append Then
                .Value = as_text
            Else
                If .Value = vbNullString Then
                    .Value = as_text
                Else
                    .Value = .Value & as_append_margin & vbLf & as_text
                End If
            End If
            .AutoSize = True
        Else
            .MultiLine = True
            .WordWrap = False ' the means to limit the width
            .AutoSize = True
            If Not as_append Then
                .Value = as_text
            Else
                If .Value = vbNullString Then
                    .Value = as_text
                Else
                    .Value = .Value & vbLf & as_text
                End If
            End If
        End If
        If as_width_limit <> 0 Then .Width = as_width_limit + 2   ' readability space
        .Height = .Height + 2 ' redability space
'        If .Parent.Height < .Top + .Height + 2 Then .Parent.Height = .Top + .Height + 2
'        If .Parent.Width < .Left + .Width Then .Parent.Width = .Left + .Width
'        If IsForm(.Parent) Then
'            With .Parent
'                If as_width_min > 0 And .Width < as_width_min Then .Width = as_width_min
'                If as_height_min > 0 And .Height < as_height_min Then .Height = as_height_min
'            End With
'        End If
    End With
        
xt: Exit Sub

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

Private Function ButtonsProvided(ByVal vButtons As Variant) As Boolean
    Select Case TypeName(vButtons)
        Case "String":                      ButtonsProvided = vButtons <> vbNullString
        Case "Collection", "Dictionary":    ButtonsProvided = vButtons.Count > 0
        Case Else:                          ButtonsProvided = True
    End Select
End Function

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

Private Sub Collect(ByRef col_into As Collection, _
                    ByVal col_with_parent As Variant, _
                    ByVal col_cntrl_type As String, _
                    ByVal col_set_height As Single, _
                    ByVal col_set_width As Single, _
           Optional ByVal col_set_visible As Boolean = False)
' ------------------------------------------------------------------------------
' Setup of a Collection (col_into) with all type (col_cntrl_type) controls
' with a parent (col_with_parent) as Collection (col_into) by assigning the
' an initial height (col_set_height) and width (col_set_width).
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim ctl As MSForms.Control
    Dim v   As Variant
    
    lBackColor = Me.BackColor
    Set col_into = New Collection

    Select Case TypeName(col_with_parent)
        Case "Collection"
            '~~ Parent is each frame in the collection
            For Each v In col_with_parent
                For Each ctl In Me.Controls
                    If TypeName(ctl) = col_cntrl_type And ctl.Parent Is v Then
                        With ctl
                            .Visible = col_set_visible
                            .Height = col_set_height
                            .Width = col_set_width
                        End With
                        col_into.Add ctl
'                        Debug.Print col_into.Count & ": " & ctl.Name
                    End If
               Next ctl
            Next v
        Case Else
            For Each ctl In Me.Controls
                If TypeName(ctl) = col_cntrl_type And ctl.Parent Is col_with_parent Then
                    With ctl
                        .Visible = col_set_visible
                        .Height = col_set_height
                        .Width = col_set_width
                    End With
                    col_into.Add ctl
'                    Debug.Print col_into.Count & ": " & ctl.Name
                End If
            Next ctl
    End Select

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub CollectDesignControls()
' ----------------------------------------------------------------------
' Collects all designed controls without considering any control's name.
' ----------------------------------------------------------------------
    Const PROC = "CollectDesignControls"
    
    On Error GoTo eh
    Dim v As Variant

    Collect col_into:=cllDsgnAreas _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=Me _
          , col_set_height:=10 _
          , col_set_width:=Me.Width - siHmarginFrames
    
    Set AreaBttns = cllDsgnAreas(2)
    Set AreaMsg = cllDsgnAreas(1)
    
    Collect col_into:=cllDsgnMsgSects _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=AreaMsg _
          , col_set_height:=50 _
          , col_set_width:=AreaMsg.Width - siHmarginFrames
    
    Collect col_into:=cllDsgnMsgSectsLabel _
          , col_cntrl_type:="Label" _
          , col_with_parent:=cllDsgnMsgSects _
          , col_set_height:=15 _
          , col_set_width:=AreaMsg.Width - (siHmarginFrames * 2)
    
    Collect col_into:=cllDsgnMsgSectsTextFrm _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=cllDsgnMsgSects _
          , col_set_height:=20 _
          , col_set_width:=AreaMsg.Width - (siHmarginFrames * 2)
    
    Collect col_into:=cllDsgnMsgSectsTextBox _
          , col_cntrl_type:="TextBox" _
          , col_with_parent:=cllDsgnMsgSectsTextFrm _
          , col_set_height:=20 _
          , col_set_width:=AreaMsg.Width - (siHmarginFrames * 3)
        
    Collect col_into:=cllDsgnBttnsFrm _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=AreaBttns _
          , col_set_height:=10 _
          , col_set_width:=10 _
          , col_set_visible:=True ' minimum is one button
    
    Collect col_into:=cllDsgnBttnRows _
          , col_cntrl_type:="Frame" _
          , col_with_parent:=cllDsgnBttnsFrm _
          , col_set_height:=10 _
          , col_set_width:=10 _
          , col_set_visible:=False ' minimum is one button
        
    Set cllDsgnBttns = New Collection
    For Each v In cllDsgnBttnRows
        Collect col_into:=cllDsgnRowBttns _
              , col_cntrl_type:="CommandButton" _
              , col_with_parent:=v _
              , col_set_height:=10 _
              , col_set_width:=siMinButtonWidth
        cllDsgnBttns.Add cllDsgnRowBttns
    Next v
    
    ProvideDictionary AppliedBttns
    ProvideDictionary AppliedBttnsRetVal

    Sections = DsgnMsgSects.Count
    Set BttnsFrm = DsgnBttnsFrm
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Function ContentHeight(ByVal ch_frame_or_form As Variant, _
                     Optional ByVal ch_visible_only As Boolean = True) As Single
' ------------------------------------------------------------------------------
' Returns the height of the controls (ch_frame_or_form) content by considering only
' applied = visible controls. When no control is provided it is the UserForm
' which is ment.
' ------------------------------------------------------------------------------
    Const PROC = "ContzentHeight"
    
    On Error GoTo eh
    Dim ctl     As MSForms.Control
    Dim i       As Long
    
    If Not IsFrameOrForm(ch_frame_or_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form - and thus has no controls!"
    
    For Each ctl In ch_frame_or_form.Controls
        With ctl
            If .Parent Is ch_frame_or_form Then
                If ch_visible_only Then
                    If ctl.Visible Then
                        ContentHeight = Max(ContentHeight, .Top + .Height)
                        i = i + 1
                    End If
                Else
                    ContentHeight = Max(ContentHeight, .Top + .Height)
                    Debug.Print ContentHeight
                    i = i + 1
                End If
            End If
        End With
    Next ctl
'    ContentHeight = ContentHeight + ScrollH_Height(ch_frame_or_form)
    If IsForm(ch_frame_or_form) Then ContentHeight = ContentHeight + 35
    
    Debug.Print i & " Controls = " & ContentHeight & " pt"

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function ContentWidth(ByVal cw_frame_or_form As Variant, _
                    Optional ByVal cw_visible_only As Boolean = True) As Single
' ------------------------------------------------------------------------------
' Returns the width of the Frame's or Form's (cw_frame_or_form) content by
' considering only visible or all controls and a possibly appliued vertical
' scrol-bar.
' ------------------------------------------------------------------------------
    Const PROC = "ContentWidth"
    
    On Error GoTo eh
    Dim ctl     As MSForms.Control
    Dim siWidth As Single
    Dim i       As Long
    
    If Not IsFrameOrForm(cw_frame_or_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form - and thus has no controls!"

    For Each ctl In cw_frame_or_form.Controls
        With ctl
            If .Parent Is cw_frame_or_form Then
                If cw_visible_only Then
                    If ctl.Visible Then
                        siWidth = Max(siWidth, (.Left + .Width))
                        i = i + 1
                    End If
                Else
                    siWidth = Max(siWidth, (.Left + .Width))
                    i = i + 1
                End If
            End If
        End With
    Next ctl
    ContentWidth = siWidth + ScrollV_Width(cw_frame_or_form)
'    Debug.Print i & " Controls = " & ContentWidth & "pt"
    If IsForm(cw_frame_or_form) Then ContentWidth = ContentWidth + 15

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

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

Private Function DsgnRowBttns(ByVal ButtonRow As Long) As Collection
' --------------------------------------------------------------------
' Return a collection of applied/use/visible buttons in row buttonrow.
' --------------------------------------------------------------------
    Set DsgnRowBttns = cllDsgnBttns(ButtonRow)
End Function

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Minimum error message display where neither mErH.ErrMsg nor mMsg.ErrMsg is
' appropriate. This is the case here because this component is used by the other
' two components which implies the danger of a loop.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
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
    If err_source = vbNullString Then err_source = Err.source
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
    ErrBttns = vbYesNoCancel
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume error line" & vbLf & _
              "No     = Resume Next (skip error line)" & vbLf & _
              "Cancel = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "fMsg." & sProc
End Function

Private Sub FrameCenterHorizontal(ByVal center_frame As MSForms.Frame, _
                         Optional ByVal within_frame As MSForms.Frame = Nothing, _
                         Optional ByVal left_margin As Single = 0)
' ------------------------------------------------------------------------------
' Center the frame (center_frame) horizontally within the frame (within_frame)
' - which defaults to the UserForm when not provided.
' ------------------------------------------------------------------------------
    
    If within_frame Is Nothing Then
        center_frame.Left = (Me.InsideWidth - center_frame.Width) / 2
    Else
        center_frame.Left = (within_frame.Width - center_frame.Width) / 2
    End If
    If center_frame.Left = 0 Then center_frame.Left = left_margin
End Sub

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

Private Sub IndicateFrameCaptionsSetup(Optional ByVal b As Boolean = True)
' ----------------------------------------------------------------------------
' When False (the default) captions are removed from all frames, else they
' remain visible for testing purpose
' ----------------------------------------------------------------------------
            
    Dim ctl As MSForms.Control
       
    If Not b Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.Caption = vbNullString
            End If
        Next ctl
    End If

End Sub

Private Function IsForm(ByVal v As Object) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = v.Parent
    IsForm = Err.Number <> 0
End Function

Private Function IsFrameOrForm(ByVal v As Object) As Boolean
    IsFrameOrForm = TypeOf v Is MSForms.UserForm Or TypeOf v Is MSForms.Frame
End Function

Private Function IsUserForm(ByVal is_obj As Object) As Boolean
      IsUserForm = TypeOf is_obj Is MSForms.UserForm
End Function

Private Function Max(ParamArray va() As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function MaxUsedWidthTextFrames() As Single
    Dim i As Long
    For i = 1 To Sections
        If IsApplied(DsgnMsgSectTxtFrm(i)) Then
            MaxUsedWidthTextFrames = Max(MaxUsedWidthTextFrames, DsgnMsgSectTxtFrm(i).Width)
        End If
    Next i
End Function

Private Function MaxWidthMsgSect(ByVal frm_area As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' The maximum usable message section width depends on the maximum message area
' width whether or not the area frame (frm_area) has a vertical scrollbar. A
' vertical scrollbar reduces the available space by the required space for the
' vertical scrollbar.
' ------------------------------------------------------------------------------
    If frm_area.ScrollBars = fmScrollBarsVertical Or frm_area.ScrollBars = fmScrollBarsBoth _
    Then MaxWidthMsgSect = MaxWidthMsgArea - SCROLL_V_WIDTH _
    Else MaxWidthMsgSect = MaxWidthMsgArea
End Function

Private Function MaxWidthMsgTextFrame(ByVal frm_area As MSForms.Frame, _
                                      ByVal frm_section As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' The maximum usable message text width depends on the maximum message section
' width and whether or not the section (frm_section) has a vertical scrollbar
' which reduces the available space by its required width.
' ------------------------------------------------------------------------------
    If frm_section.ScrollBars = fmScrollBarsVertical Or frm_section.ScrollBars = fmScrollBarsBoth _
    Then MaxWidthMsgTextFrame = MaxWidthMsgSect(frm_area) - SCROLL_V_WIDTH _
    Else MaxWidthMsgTextFrame = MaxWidthMsgSect(frm_area)
End Function

Private Function MaxWidthSectTxtBox(ByVal frm_text As MSForms.Frame) As Single
' ------------------------------------------------------------------------------
' The maximum with of a sections text-box depends on whether or not the frame of
' the TextBox (frm_text) has a vertical scrollbar which reduces the available
' space by its required width.
' ------------------------------------------------------------------------------
    If frm_text.ScrollBars = fmScrollBarsVertical Or frm_text.ScrollBars = fmScrollBarsBoth _
    Then MaxWidthSectTxtBox = frm_text.Width - SCROLL_V_WIDTH _
    Else MaxWidthSectTxtBox = frm_text.Width
End Function

Private Function Min(ParamArray va() As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ----------------------------------------------------------------------------
    Dim v   As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Private Property Get MonitorHeightExSteps() As Single
    MonitorHeightExSteps = ContentHeight(frmSteps.Parent) - frmSteps.Height
End Property

Private Property Get MonitorHeightMaxSteps()
    MonitorHeightMaxSteps = Me.MsgHeightMax - MonitorHeightExSteps
End Property

Private Sub MonitorEstablishFooter(ByVal mf_top As Single)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------

    If tbxFooter Is Nothing Then
        Set tbxFooter = AddControl(ac_ctl:=TextBox, ac_visible:=True)
        SetupTextFont tbxFooter, m_footer
        With tbxFooter
            .Top = mf_top + 6
            .Left = 0
            .Value = TextMonitorFooter.Text
            .Height = 18
            .Width = Me.InsideWidth
            .BackColor = Me.BackColor
            .BorderColor = Me.BackColor
            .BorderStyle = fmBorderStyleSingle
            Me.Height = .Top + .Height + 30 + ScrollH_Height(frmSteps)
        End With
    Else
        tbxFooter.Value = TextMonitorFooter.Text
    End If

End Sub

Private Sub SetupTextFont(ByVal ctl As MSForms.Control, _
                          ByVal kind_of_text As KindOfText)
' ------------------------------------------------------------------------------
' Setup the font properties for a Label or TextBox (ctl) according to the
' corresponding TypeMsgText type (kind_of_text).
' ------------------------------------------------------------------------------

    Dim txt As TypeMsgText
    txt = Me.Text(kind_of_text)
    
    With ctl.Font
        If .Bold <> txt.FontBold Then .Bold = txt.FontBold
        If .Italic <> txt.FontItalic Then .Italic = txt.FontItalic
        If .Underline <> txt.FontUnderline Then .Underline = txt.FontUnderline
        If txt.MonoSpaced Then
            .name = DFLT_TXT_MONOSPACED_FONT_NAME
            If txt.FontSize = 0 _
            Then .Size = DFLT_TXT_MONOSPACED_FONT_SIZE _
            Else .Size = txt.FontSize
        Else
            If txt.FontName = vbNullString _
            Then .name = DFLT_TXT_PROPSPACED_FONT_NAME _
            Else .name = txt.FontName
            If txt.FontSize = 0 _
            Then .Size = DFLT_TXT_PROPSPACED_FONT_SIZE _
            Else .Size = txt.FontSize
        End If
    End With
    ctl.ForeColor = txt.FontColor
    If Me.VisualizeControls Then ctl.BackColor = CTL_VISUALIZE_TBX_BACKCOL
End Sub

Private Sub MonitorEstablishHeader(ByRef mh_top As Single)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    
    If tbxHeader Is Nothing Then
        Set tbxHeader = AddControl(ac_ctl:=TextBox, ac_visible:=False)
        SetupTextFont tbxHeader, m_header
        With tbxHeader
            .Top = mh_top
            .Left = 0
            .Value = TextMonitorHeader.Text
            .Height = 18
            .Width = Me.InsideWidth
            .BackColor = Me.BackColor
            .BorderColor = Me.BackColor
            .BorderStyle = fmBorderStyleSingle
            .Visible = True
            mh_top = AdjustToVgrid(.Top + .Height)
        End With
    Else
        tbxHeader.Value = TextMonitorHeader.Text
    End If

End Sub

Private Sub MonitorEstablishStep(ByRef ms_top As Single, _
                        Optional ByVal ms_in As MSForms.Frame = Nothing)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Set tbxStep = AddControl(ac_ctl:=TextBox, ac_in:=ms_in, ac_visible:=False)
    SetupTextFont tbxStep, m_step
    With tbxStep
        .Top = ms_top
        .Left = 0
        .Visible = True
        .Height = 18
        .Width = Me.InsideWidth
        .BorderColor = Me.BackColor
        .BorderStyle = fmBorderStyleSingle
         ms_top = AdjustToVgrid(.Top + .Height)
    End With

End Sub

Public Sub MonitorInitialize(ByVal mon_title As String, _
                             ByVal mon_steps_displayed As Long)
' ------------------------------------------------------------------------------
' Setup header (mon_header), steps to be displayed (mon_steps_displayed) - at
' first invisible - and the footer (mon_footer).
' ------------------------------------------------------------------------------
    Const PROC = "MonitorInitialize"
    
    On Error GoTo eh
    Dim ctl                 As MSForms.Control
    Dim siTop               As Single
    Dim i                   As Long
    Dim tbx                 As MSForms.TextBox
    Dim siStepsHeightMax    As Single
    Dim siNetHeight         As Single           ' The height of the setup header and footer
    
    If Not bMonitorInitialized Then
        lMonitorSteps = mon_steps_displayed
        For Each ctl In Me.Controls
            ctl.Visible = False
        Next ctl
        siTop = 6
        
        With Me
            .Caption = mon_title
            .Width = .MsgWidthMin
            
            '~~ Establish monitor header
            If TextMonitorHeader.Text <> vbNullString _
            Then MonitorEstablishHeader mh_top:=siTop
            With tbxHeader
            End With
            
            '~~ Establish monitor steps in a dedicated frame
            Set frmSteps = AddControl(ac_ctl:=Frame, ac_name:="frmAreaSteps")
            With frmSteps
                .Top = siTop
                .Visible = True
                .BorderColor = Me.BackColor
                .BorderStyle = fmBorderStyleSingle
                If Me.VisualizeControls _
                Then .BackColor = CTL_VISUALIZE_TEXT_SECTION _
                Else .BackColor = Me.BackColor

                siTop = 0
                For i = 1 To lMonitorSteps
                    MonitorEstablishStep ms_in:=frmSteps, ms_top:=siTop
                    Qenqueue cllSteps, tbxStep
                Next i
                .Height = ContentHeight(frmSteps, False)
                '~~ The maximum height for the steps frame is the max formheight minus the height of header and footer
                siNetHeight = ContentHeight(frmSteps.Parent) - frmSteps.Height
                siStepsHeightMax = Me.MsgHeightMax - siNetHeight
                NewHeight(frmSteps, False) = Min(siStepsHeightMax, .Height)
                '~~ Establish monitor footer
                If TextMonitorFooter.Text <> vbNullString Then
                    siTop = AdjustToVgrid(frmSteps.Top + .Height)
                    MonitorEstablishFooter mf_top:=siTop + 6
                End If
            End With
            NewHeight(frmSteps.Parent) = Min(.MsgHeightMax, ContentHeight(frmSteps.Parent))
        End With
        bMonitorInitialized = True
    End If

'    Debug.Print "Me.MsgHeightMax = " & Me.MsgHeightMax & " (" & mMsg.Prcnt(Me.MsgHeightMax, "H") & "%)"
'    Debug.Print "Me.MsgWidthMax  = " & Me.MsgWidthMax & " (" & mMsg.Prcnt(Me.MsgWidthMax, "W") & "%)"


xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub MonitorStep()
' ------------------------------------------------------------------------------
' Display a step. Note that the height of the steps frame (frmSteps) is already
' adjusted to the number of steps to be displayed. However, when one or another
' step's height is more than one line the height needs to be ajusted.
' ------------------------------------------------------------------------------
    Const PROC = "MonitirStep"
    
    On Error GoTo eh
    Dim siMaxWidth          As Single
    Dim tbx                 As MSForms.TextBox
    Dim i                   As Long
    Dim siTop               As Single
    Dim siNetHeight         As Single
    Dim siStepsHeightMax    As Single
    
    siTop = 0
    If TextMonitorStep.Text <> vbNullString Then
        If lStepsDisplayed < lMonitorSteps Then
            If lStepsDisplayed > 0 Then
                Set tbx = cllSteps(lStepsDisplayed)
                siTop = AdjustToVgrid(tbx.Top + tbx.Height)
            End If
            Set tbx = cllSteps(lStepsDisplayed + 1)
            SetupTextFont tbx, m_step
            tbx.Visible = True
            tbx.Top = siTop
            
            If TextMonitorStep.MonoSpaced Then
                AutoSizeTextBox as_tbx:=tbx _
                              , as_text:=TextMonitorStep.Text _
                              , as_width_limit:=0
            Else
                AutoSizeTextBox as_tbx:=tbx _
                              , as_width_limit:=Me.InsideWidth _
                              , as_text:=TextMonitorStep.Text
            End If
            MonitorStepsAdjustTopPosition
            NewWidth(frmSteps, False) = Min(Me.MsgWidthMax, ContentWidth(frmSteps, False)) ' applies a horizontal scroll-bar or adjust its width
            NewWidth(frmSteps.Parent) = ContentWidth(frmSteps.Parent)
            
            siNetHeight = Me.Height - (frmSteps.Height - frmSteps.Top)
            siStepsHeightMax = Me.MsgHeightMax - siNetHeight
            NewHeight(frmSteps, False, fmScrollActionBegin) = Min(MonitorHeightMaxSteps, ContentHeight(frmSteps, False) + ScrollH_Height(frmSteps))
            
            lStepsDisplayed = lStepsDisplayed + 1
            If Not tbxFooter Is Nothing Then
                siTop = AdjustToVgrid(frmSteps.Top + frmSteps.Height + 6)
                If siTop > tbxFooter.Top Then
                    tbxFooter.Top = siTop
                    Me.Height = ContentHeight(frmSteps.Parent)
                End If
            End If
        
        Else
            '~~ All steps are displayed
            siTop = 0
            Set tbx = Qdequeue(cllSteps)
            tbx.Value = vbNullString
                        
            Qenqueue cllSteps, tbx
            
            For i = 1 To lMonitorSteps
                Set tbx = cllSteps(i)
                tbx.Top = siTop
                siTop = AdjustToVgrid(tbx.Top + tbx.Height)
                siMaxWidth = Max(siMaxWidth, tbx.Width)
            Next i
            
            If TextMonitorStep.MonoSpaced Then
                AutoSizeTextBox as_tbx:=tbx _
                              , as_text:=TextMonitorStep.Text _
                              , as_width_limit:=0
            Else
                AutoSizeTextBox as_tbx:=tbx _
                              , as_width_limit:=Me.InsideWidth _
                              , as_text:=TextMonitorStep.Text
            End If
            MonitorStepsAdjustTopPosition
            NewWidth(frmSteps, False) = Min(Me.MsgWidthMax, ContentWidth(frmSteps, False)) ' applies a horizontal scroll-bar or adjust its width
            NewWidth(frmSteps.Parent) = ContentWidth(frmSteps.Parent)
            
            siNetHeight = Me.Height - (frmSteps.Height - frmSteps.Top)
            siStepsHeightMax = Me.MsgHeightMax - siNetHeight
            NewHeight(frmSteps, False, fmScrollActionEnd) = Min(MonitorHeightMaxSteps, ContentHeight(frmSteps, False) + ScrollH_Height(frmSteps))
        End If
    End If
    
    If TextMonitorFooter.Text <> vbNullString Then
        siTop = AdjustToVgrid(frmSteps.Top + frmSteps.Height + 6)
        MonitorEstablishFooter mf_top:=siTop
    End If
    
    DoEvents
    Me.Height = ContentHeight(frmSteps.Parent)
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub PositionMessageOnScreen( _
           Optional ByVal pos_top_left As Boolean = False)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    
    On Error Resume Next
        
    With Me
        .StartUpPosition = sup_Manual
        If pos_top_left Then
            .Left = 5
            .Top = 5
        Else
            .Left = (VirtualScreenWidthPts - .Width) / 2
            .Top = (VirtualScreenHeightPts - .Height) / 4
        End If
    End With
    
    '~~ First make sure the bottom right fits,
    '~~ then check if the top-left is still on the screen (which gets priority).
    With Me
        If ((.Left + .Width) > (VirtualScreenLeftPts + VirtualScreenWidthPts)) Then .Left = ((VirtualScreenLeftPts + VirtualScreenWidthPts) - .Width)
        If ((.Top + .Height) > (VirtualScreenTopPts + VirtualScreenHeightPts)) Then .Top = ((VirtualScreenTopPts + VirtualScreenHeightPts) - .Height)
        If (.Left < VirtualScreenLeftPts) Then .Left = VirtualScreenLeftPts
        If (.Top < VirtualScreenTopPts) Then .Top = VirtualScreenTopPts
    End With
    
End Sub

Private Sub ProvideDictionary(ByRef dct As Dictionary)
' ----------------------------------------------------
' Provides a clean or new Dictionary.
' ----------------------------------------------------
    If Not dct Is Nothing Then dct.RemoveAll Else Set dct = New Dictionary
End Sub

Private Function Qdequeue(ByRef qu As Collection) As Variant
    Const PROC = "DeQueue"
    
    On Error GoTo eh
    If qu Is Nothing Then GoTo xt
    If QisEmpty(qu) Then GoTo xt
    On Error Resume Next
    Set Qdequeue = qu(1)
    If Err.Number <> 0 _
    Then Qdequeue = qu(1)
    qu.Remove 1

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub Qenqueue(ByRef qu As Collection, ByVal qu_item As Variant)
    If qu Is Nothing Then Set qu = New Collection
    qu.Add qu_item
End Sub

Private Function QisEmpty(ByVal qu As Collection) As Boolean
    If Not qu Is Nothing _
    Then QisEmpty = qu.Count = 0 _
    Else QisEmpty = True
End Function

Private Function ScrollH_Applied(ByVal sa_frame_or_form As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the control (sa_frame_or_form) has a horizontal scrollbar applied. When
' no control is provided it is the UserForm which is ment.
' ------------------------------------------------------------------------------
    If IsFrameOrForm(sa_frame_or_form) Then
        Select Case sa_frame_or_form.ScrollBars
            Case fmScrollBarsBoth, fmScrollBarsHorizontal: ScrollH_Applied = True
        End Select
    End If
End Function

Private Sub ScrollH_Apply(ByVal sa_width As Single, _
                          ByRef sa_frame_or_form As Variant, _
                 Optional ByVal sa_x_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' Apply a horizontal scrollbar is applied to the frame (sa_frame_or_form) and
' adjusted to the frame content's width (sa_width). In case a horizontal
' scrollbar is already applied only its width is adjusted.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollH_Apply"
    
    On Error GoTo eh
    If Not IsFrameOrForm(sa_frame_or_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
    
    With sa_frame_or_form
        Select Case .ScrollBars
            Case fmScrollBarsBoth
                '~~ The already displayed horizonzal scrollbar's width is adjusted
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollWidth = sa_width
                .Scroll xAction:=sa_x_action
            Case fmScrollBarsHorizontal
                '~~ Already displayed (no vertical scrollbar yet)
                '~~ No need to adjust the height for the scrollbar
                .KeepScrollBarsVisible = fmScrollBarsHorizontal
                .ScrollWidth = sa_width
                .Scroll xAction:=sa_x_action
                .Height = ContentHeight(sa_frame_or_form)
            Case fmScrollBarsVertical
                '~~ Add a horizontal scrollbar to the already displayed vertical
                .ScrollBars = fmScrollBarsBoth
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollWidth = sa_width
                .Scroll xAction:=sa_x_action
            Case fmScrollBarsNone
                '~~ Add a horizontal scrollbar
                .ScrollBars = fmScrollBarsHorizontal
                .KeepScrollBarsVisible = fmScrollBarsHorizontal
                .ScrollWidth = sa_width
                .Scroll xAction:=sa_x_action
                .Height = ContentHeight(sa_frame_or_form)
        End Select
    End With

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function ScrollH_Height(ByVal sh_frame_or_form As Variant) As Single
    If IsFrameOrForm(sh_frame_or_form) Then
        If ScrollH_Applied(sh_frame_or_form) Then ScrollH_Height = SCROLL_H_HEIGHT
    End If
End Function

Private Function ScrollV_Applied(Optional ByVal sa_frame_or_form As Variant = Nothing) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the control (sa_frame_or_form) has a vertical scrollbar applied. When no
' control is provided it is the UserForm which is ment.
' ------------------------------------------------------------------------------
    If IsFrameOrForm(sa_frame_or_form) Then
        Select Case sa_frame_or_form.ScrollBars
            Case fmScrollBarsBoth, fmScrollBarsVertical: ScrollV_Applied = True
        End Select
    End If
End Function

Private Sub ScrollV_Apply(ByVal sa_height As Single, _
                          ByRef sa_frame_or_form As Variant, _
                 Optional ByVal sa_y_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' A vertical scrollbar is applied to the frame (sa_frame_or_form) and adjusted to
' the frame content's height (sa_height). In case a vertical scrollbar is
' already applied only its width is adjusted.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollV_Apply"
    
    On Error GoTo eh
    If Not IsFrameOrForm(sa_frame_or_form) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a Frame nor a Form!"
    
    With sa_frame_or_form
        Select Case .ScrollBars
            Case fmScrollBarsBoth
                '~~ The already displayed horizonzal scrollbar's width is adjusted
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollHeight = sa_height
                .Scroll yAction:=sa_y_action
            Case fmScrollBarsHorizontal
                '~~ Already displayed (no vertical scrollbar yet)
                '~~ No need to adjust the height for the scrollbar
                .ScrollBars = fmScrollBarsBoth
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollHeight = sa_height
                .Scroll yAction:=sa_y_action
            Case fmScrollBarsVertical
                '~~ Add a horizontal scrollbar to the already displayed vertical
                .KeepScrollBarsVisible = fmScrollBarsVertical
                .ScrollHeight = sa_height
                .Scroll yAction:=sa_y_action
            Case fmScrollBarsNone
                '~~ Add a horizontal scrollbar
                .ScrollBars = fmScrollBarsVertical
                .KeepScrollBarsVisible = fmScrollBarsVertical
                .ScrollHeight = sa_height
                .Scroll yAction:=sa_y_action
        End Select
        .Width = ContentWidth(sa_frame_or_form)
    End With
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ScrollV_MsgSectionOrArea(ByVal exceeding_height As Single)
' ------------------------------------------------------------------------------
' Either because the message area occupies 60% or more of the total height or
' because both, the message area and the buttons area us about the same height,
' it - or only the section text occupying 65% or more - will be reduced by the
' exceeding height amount (exceeding_height) and will get a vertical scrollbar.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollV_MsgSectionOrArea"
    
    On Error GoTo eh
    Dim i                   As Long
    Dim VScrollApplied      As Boolean
    
    '~~ Find a/the message section text which occupies 65% or more of the message area's height,
    For i = 1 To Sections
        SetCtrlsOfSection i
        
        If SectTxtFrm.Height >= AreaMsg.Height * 0.65 _
        Or ScrollV_Applied(SectTxtFrm) Then
            ' ------------------------------------------------------------------------------
            ' There is a section which occupies 65% of the overall height or has already a
            ' vertical scrollbar applied. Assigning a new frame height applies a vertical
            ' scrollbar if none is applied yet or just adjusts the scrollbar's height to the
            ' frame's content height
            ' ------------------------------------------------------------------------------
            If UsageType = usage_progress_display Then
'                Debug.Print SectTxtFrm.Height - exceeding_height
                NewHeight(SectTxtFrm, fmScrollActionEnd) = SectTxtFrm.Height - exceeding_height
                AdjustedParentsWidthAndHeight SectTxtBox
                AdjustTopPositions
                SetCtrlsOfSection i ' reset
                VScrollApplied = ScrollV_Applied(SectTxtFrm)
                Exit For
            Else
                If SectTxtFrm.Height - exceeding_height > 0 Then
'                    Debug.Print SectTxtFrm.Height - exceeding_height
                    NewHeight(SectTxtFrm) = SectTxtFrm.Height - exceeding_height
                    AdjustedParentsWidthAndHeight SectTxtBox
                    AdjustTopPositions
                    SetCtrlsOfSection i ' reset
                    VScrollApplied = ScrollV_Applied(SectTxtFrm)
                    Exit For
                End If
            End If
        End If
    Next i
    
    If Not VScrollApplied Then
        '~~ None of the message sections has a dominating height. Becaue the overall message area
        '~~ occupies >=60% of the height it is now reduced to fit the maximum message height
        '~~ thereby receiving a vertical scroll-bar
        NewHeight(AreaMsg) = ContentHeight(AreaMsg) - exceeding_height
        AdjustedParentsWidthAndHeight SectTxtBox
        AdjustTopPositions
    End If

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ScrollV_WhereApplicable()
' ------------------------------------------------------------------------------
' Reduce the height of the message area and or the height of the buttons area to
' have the message form not exceeds the specified maximum height. The area which
' uses 60% or more of the overall height is the one being reduced. Otherwise
' both are reduced proportionally.
' When one of the message sections within the to be reduced message area
' occupies 80% or more of the overall message area height only this section
' is reduced and gets a verticall scrollbar.
' The reduced frames are returned (frame_msg, frame_bttns).
' ------------------------------------------------------------------------------
    Const PROC = "ScrollV_WhereApplicable"
    
    On Error GoTo eh
    Dim TotalExceedingHeight    As Single
    
    '~~ When the message form's height exceeds the specified maximum height
    If Me.Height > siMsgHeightMax Then
        With Me
            TotalExceedingHeight = .Height - siMsgHeightMax
            If TotalExceedingHeight < 20 Then GoTo xt ' not worth any intervention
            .Height = siMsgHeightMax     '~~ Reduce the height to the max height specified
            
            If PrcntgHeightMsgArea >= 0.6 Then
                '~~ Either the message area as a whole or the dominating message section - if theres is any -
                '~~ will be height reduced and applied with a vertical scroll bar
                ScrollV_MsgSectionOrArea TotalExceedingHeight
            ElseIf PrcntgHeightAreaBttns >= 0.6 Then
                '~~ Only the buttons area will be reduced and applied with a vertical scrollbar.
'                Debug.Print AreaBttns.Height - TotalExceedingHeight
                NewHeight(AreaBttns) = AreaBttns.Height - TotalExceedingHeight
            Else
                '~~ Both, the message area and the buttons area will be
                '~~ height reduced proportionally and applied with a vertical scrollbar
'                Debug.Print AreaMsg.Height - (TotalExceedingHeight * PrcntgHeightMsgArea)
                NewHeight(AreaMsg) = AreaMsg.Height - (TotalExceedingHeight * PrcntgHeightMsgArea)
'                Debug.Print AreaBttns.Height - (TotalExceedingHeight * PrcntgHeightAreaBttns)
                NewHeight(AreaBttns) = AreaBttns.Height - (TotalExceedingHeight * PrcntgHeightAreaBttns)
            End If
        End With
    End If ' height exceeds specified maximum
   
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function ScrollV_Width(ByVal sw_frame_or_form As Variant) As Single
    If IsFrameOrForm(sw_frame_or_form) Then
        If ScrollV_Applied(sw_frame_or_form) Then ScrollV_Width = SCROLL_V_WIDTH
    End If
End Function

Private Sub SetCtrlsOfSection(ByVal i As Long)
    Set SectFrm = DsgnMsgSect(i)
    Set SectLbl = DsgnMsgSectLbl(i)
    Set SectTxtFrm = DsgnMsgSectTxtFrm(i)
    Set SectTxtBox = DsgnMsgSectTxtBox(i)
    MsgSectTxt = MsgText(i)
    MsgSectLbl = MsgLabel(i)
End Sub

Public Sub Setup()
    Const PROC = "Setup"
    
    On Error GoTo eh
    
    VisualizeControlsSetup
    CollectDesignControls
            
    IndicateFrameCaptionsSetup bIndicateFrameCaptions ' may be True for test purpose
    
    '~~ Start the setup as if there wouldn't be any message - which might be the case
    Me.StartUpPosition = 2
    Me.Height = 200                             ' just to start with - specifically for test purpose
    Me.Width = siMsgWidthMin
    
'    PositionMessageOnScreen pos_top_left:=True  ' in case of test best pos to start with
    AreaMsg.Visible = False
    AreaBttns.Top = VSPACE_AREAS
    
    '~~ ----------------------------------------------------------------------------------------
    '~~ The  p r i m a r y  setup of the title, the message sections and the reply buttons
    '~~ returns their individual widths which determines the minimum required message form width
    '~~ This setup ends width the final message form width and all elements adjusted to it.
    '~~ ----------------------------------------------------------------------------------------
    
    '~~ Setup of the title, the first element which potentially effects the final message width
    If Not bDoneTitle _
    Then Setup1_Title setup_title:=sMsgTitle _
                    , setup_width_min:=siMsgWidthMin _
                    , setup_width_max:=siMsgWidthMax
    
    '~~ Setup of any monospaced message sections, the second element which potentially effects the final message width.
    '~~ In case the section width exceeds the maximum width specified a horizontal scrollbar is applied.
    Setup2_MsgSectsMonoSpaced
    
    '~~ Setup the reply buttons. This is the third element which may effect the final message's width.
    '~~ In case the widest buttons row exceeds the maximum width specified for the message
    '~~ a horizontal scrollbar is applied.
    If ButtonsProvided(vButtons) Then
        Setup3_Bttns vButtons
        SizeAndPosition2Bttns1
        SizeAndPosition2Bttns2Rows
        SizeAndPosition2Bttns3Frame
        AdjustedParentsWidthAndHeight AreaBttns
        SizeAndPosition2Bttns4Area
    End If
    
    ' -----------------------------------------------------------------------------------------------
    ' At this point the form has reached its final width (all proportionally spaced message sections
    ' are adjusted to it). However, the message height is only final in case there are just buttons
    ' but no message. The setup of proportional spaced message sections determines the final message
    ' height. When it exeeds the maximum height specified one or two vertical scrollbars are applied.
    ' -----------------------------------------------------------------------------------------------
    Setup4_MsgSectsPropSpaced
                
    ' -----------------------------------------------------------------------------------------------
    ' When the message form height exceeds the specified or the default message height the height of
    ' the message area and or the buttons area is reduced and a vertical is applied.
    ' When both areas are about the same height (neither is taller the than 60% of the total heigth)
    ' both will get a vertical scrollbar, else only the one which uses 60% or more of the height.
    ' -----------------------------------------------------------------------------------------------
    AdjustTopPositions
    ScrollV_WhereApplicable
    
    '~~ Final form width adjustment
    '~~ When the message area or the buttons area has a vertical scrollbar applied
    '~~ the scrollbar may not be visible when the width as a result exeeds the specified
    '~~ message form width. In order not to interfere again with the width of all content
    '~~ the message form width is extended (over the specified maximum) in order to have
    '~~ the vertical scrollbar visible
    AdjustTopPositions
    PositionMessageOnScreen
    bSetUpDone = True ' To indicate for the Activate event that the setup had already be done beforehand
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Setup1_Title( _
                ByVal setup_title As String, _
                ByVal setup_width_min As Single, _
                ByVal setup_width_max As Single)
' ------------------------------------------------------------------------------
' Setup the message form for the provided title (setup_title) optimized with the
' provided minimum width (setup_width_min) and the provided maximum width
' (setup_width_max) by using a certain factor (setup_factor) for the calculation
' of the width required to display an untruncated title - as long as the maximum
' widht is not exeeded.
' ------------------------------------------------------------------------------
    Const PROC = "Setup1_Title"
    Const FACTOR = 1.45
    
    On Error GoTo eh
    Dim Correction    As Single
    
    With Me
        .Width = setup_width_min
        '~~ The extra title label is only used to adjust the form width and remains hidden
        With .laMsgTitle
            With .Font
                .Bold = False
                .name = Me.Font.name
                .Size = 8    ' Value which comes to a length close to the length required
            End With
            .Caption = vbNullString
            .AutoSize = True
            .Caption = " " & setup_title    ' some left margin
        End With
        .Caption = setup_title
        Correction = (CInt(.laMsgTitle.Width)) / 2000
        .Width = Min(setup_width_max, .laMsgTitle.Width * (FACTOR - Correction))
        .Width = Max(.Width, setup_width_min)
        TitleWidth = .Width
    End With
    bDoneTitle = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup2_MsgSectsMonoSpaced()
' --------------------------------------------------------------------------------------
' Setup of all sections for which a text is provided indicated mono-spaced.
' --------------------------------------------------------------------------------------
    Const PROC = "Setup2_MsgSectsMonoSpaced"
    
    On Error GoTo eh
    Dim i As Long
    
    For i = 1 To Sections
        With Me.MsgText(i)
            If .MonoSpaced And .Text <> vbNullString Then
                SetCtrlsOfSection i
                SetupMsgSect
                AdjustedParentsWidthAndHeight SectTxtBox
                AdjustTopPositions
                AdjustedParentsWidthAndHeight AreaMsg
            End If
        End With
    Next i
    bDoneMonoSpacedSects = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup3_Bttns(ByVal vButtons As Variant)
' --------------------------------------------------------------------------------------
' Setup the reply buttons (vButtons) provided.
' Note: When the provided vButtons argument is a string it wil be converted into a
'       collection and the procedure is performed recursively with it.
' --------------------------------------------------------------------------------------
    Const PROC = "Setup3_Bttns"
    
    On Error GoTo eh
    
    '~~ Setup all reply button by calculatig their maximum width and height
    Select Case TypeName(vButtons)
        Case "Long":        SetupBttnsFromValue vButtons ' buttons are specified by one single VBA.MsgBox button value only
        Case "String":      SetupBttnsFromString vButtons
        Case "Collection":  SetupBttnsFromCollection vButtons
        Case "Dictionary":  SetupBttnsFromCollection vButtons
        Case Else
            '~~ Because vbuttons is not provided by a known/accepted format
            '~~ the message will be setup with an Ok only button", vbExclamation
            Setup3_Bttns vbOKOnly
    End Select
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub Setup4_MsgSectsPropSpaced()
    Const PROC = "Setup4_MsgSectsPropSpaced"
    
    On Error GoTo eh
    Dim i As Long
    
    For i = 1 To Sections
        SetCtrlsOfSection i
        If MsgSectTxt.Text <> vbNullString And Not MsgSectTxt.MonoSpaced Then
            SetupMsgSect
        End If
    Next i
    bDonePropSpacedSects = True
    bDoneMsgArea = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupBttnsFromCollection(ByVal cllButtons As Collection)
' -------------------------------------------------------------------------------
' Setup the reply buttons based on the comma delimited string of button captions
' and row breaks indicated by a vbLf, vbCr, or vbCrLf.
' ---------------------------------------------------------------------
    Const PROC = "SetupBttnsFromCollection"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim BttnRow     As MSForms.Frame
    Dim Bttn        As MSForms.CommandButton

    If cllButtons.Count = 0 Then GoTo xt
    IsApplied(AreaBttns) = True
    IsApplied(DsgnBttnsFrm) = True

    lSetupRows = 1
    lSetupRowButtons = 0
    Set BttnRow = DsgnBttnRow(1)
    Set Bttn = DsgnBttn(1, 1)
    
    Me.Height = 100 ' just to start with
    AreaBttns.Top = VSPACE_AREAS
    BttnsFrm.Top = AreaBttns.Top
    BttnRow.Top = BttnsFrm.Top
    Bttn.Top = BttnRow.Top
    Bttn.Width = DFLT_BTTN_MIN_WIDTH
    
    For Each v In cllButtons
        If IsNumeric(v) Then v = mMsg.BttnsArgs(v)
        Select Case v
            Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore, vbYesNo, vbResumeOk
                SetupBttnsFromValue v
            Case Else
                If v <> vbNullString Then
                    If v = vbLf Or v = vbCr Or v = vbCrLf Then
                        '~~ prepare for the next row
                        If lSetupRows <= 7 Then ' ignore exceeding rows
                            IsApplied(DsgnBttnRow(lSetupRows)) = True
                            lSetupRows = lSetupRows + 1
                            lSetupRowButtons = 0
                        Else
                            MsgBox "Setup of button row " & lSetupRows & " ignored! The maximum applicable rows is 7."
                        End If
                    Else
                        lSetupRowButtons = lSetupRowButtons + 1
                        If lSetupRowButtons <= 7 And lSetupRows <= 7 Then
                            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:=v, sb_ret_value:=v
                        Else
                            MsgBox "The setup of button " & lSetupRowButtons & " in row " & lSetupRows & " is ignored! The maximum applicable buttons per row is 7 " & _
                                   "and the maximum rows is 7 !"
                        End If
                    End If
                End If
        End Select
    Next v
    If lSetupRows <= 7 Then
        IsApplied(DsgnBttnRow(lSetupRows)) = True
    End If
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupBttnsFromString(ByVal buttons_string As String)
    
    Dim cll As New Collection
    Dim v   As Variant
    
    For Each v In Split(buttons_string, ",")
        cll.Add v
    Next v
    Setup3_Bttns cll
    
End Sub

Private Sub SetupBttnsFromValue(ByVal lButtons As Long)
' -------------------------------------------------------------------------------
' Setup a row of standard VB MsgBox reply command buttons
' -------------------------------------------------------------------------------
    Const PROC = "SetupBttnsFromValue"
    
    On Error GoTo eh
    Dim ResumeErrorLine As String: ResumeErrorLine = "Resume" & vbLf & "Error Line"
    Dim PassOn          As String: PassOn = "Pass on Error to" & vbLf & "Entry Procedure"
    
    Select Case lButtons
        Case vbOKOnly
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
        Case vbOKCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Cancel", sb_ret_value:=vbCancel
        Case vbYesNo
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Yes", sb_ret_value:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="No", sb_ret_value:=vbNo
        Case vbRetryCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Retry", sb_ret_value:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Cancel", sb_ret_value:=vbCancel
        Case vbResumeOk
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:=ResumeErrorLine, sb_ret_value:=vbResume
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
        Case vbYesNoCancel
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Yes", sb_ret_value:=vbYes
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="No", sb_ret_value:=vbNo
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Cancel", sb_ret_value:=vbCancel
        Case vbAbortRetryIgnore
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Abort", sb_ret_value:=vbAbort
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Retry", sb_ret_value:=vbRetry
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Ignore", sb_ret_value:=vbIgnore
        Case vbResumeOk
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Resume" & vbLf & "Error Line", sb_ret_value:=vbResume
            lSetupRowButtons = lSetupRowButtons + 1
            SetupButton ButtonRow:=lSetupRows, sb_index:=lSetupRowButtons, sb_caption:="Ok", sb_ret_value:=vbOK
    
        Case Else
            MsgBox "The value provided for the ""buttons"" argument is not a known VB MsgBox value"
    End Select
    If lSetupRows <> 0 Then
        IsApplied(DsgnBttnRow(lSetupRows)) = True
        IsApplied(DsgnBttnsFrm) = True
    End If
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupButton(ByVal ButtonRow As Long, _
                        ByVal sb_index As Long, _
                        ByVal sb_caption As String, _
                        ByVal sb_ret_value As Variant)
' -------------------------------------------------------------------------------
' Setup an applied reply sb_index's (sb_index) visibility and caption,
' calculate the maximum sb_index width and height, keep a record of the setup
' reply sb_index's return value.
' -------------------------------------------------------------------------------
    Const PROC = "SetupButton"
    
    On Error GoTo eh
    Dim cmb As MSForms.CommandButton
    
    If ButtonRow = 0 Then ButtonRow = 1
    Set cmb = DsgnBttn(ButtonRow, sb_index)
    
    With cmb
        .AutoSize = True
        .WordWrap = False ' the longest line determines the sb_index's width
        .Caption = sb_caption
        .AutoSize = False
        .Height = .Height + 1 ' safety margin to ensure proper multilin caption display
        siMaxButtonHeight = Max(siMaxButtonHeight, .Height)
        siMaxButtonWidth = Max(siMaxButtonWidth, .Width, siMinButtonWidth)
    End With
    AppliedBttns.Add cmb, ButtonRow
    AppliedButtonRetVal(cmb) = sb_ret_value ' keep record of the setup sb_index's reply value
    IsApplied(cmb) = True
    IsApplied(DsgnBttnRow(ButtonRow)) = True
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupMsgSect()
' -------------------------------------------------------------------------------
' Setup a message section with its label when one is specified and return the
' Message's width when greater than any other.
' Note: All height adjustments except the one for the text box are done by the
'       SizeAndPosition
' -------------------------------------------------------------------------------
    Const PROC = "SetupMsgSect"
    
    On Error GoTo eh
                
    SectFrm.Width = AreaMsg.Width
    SectLbl.Width = SectFrm.Width
    SectTxtFrm.Width = SectFrm.Width
    SectTxtBox.Width = SectFrm.Width
        
    If MsgSectTxt.Text <> vbNullString Then
    
        IsApplied(AreaMsg) = True
        IsApplied(SectFrm) = True
        IsApplied(SectTxtFrm) = True
        IsApplied(SectTxtBox) = True
'        Debug.Print SectFrm.Name
'        Debug.Print SectTxtFrm.Name
'        Debug.Print SectTxtBox.Name
                
        If MsgSectLbl.Text <> vbNullString Then
            IsApplied(SectLbl) = True
'            Debug.Print SectLbl.Name
            With SectLbl
                .Left = 10
                .Width = Me.InsideWidth - (siHmarginFrames * 2)
                .Caption = MsgSectLbl.Text
                With .Font
                    If MsgSectLbl.MonoSpaced Then
                        If MsgSectLbl.FontName <> vbNullString Then .name = MsgSectLbl.FontName Else .name = DFLT_LBL_MONOSPACED_FONT_NAME
                        If MsgSectLbl.FontSize <> 0 Then .Size = MsgSectLbl.FontSize Else .Size = DFLT_LBL_MONOSPACED_FONT_SIZE
                    Else
                        If MsgSectLbl.FontName <> vbNullString Then .name = MsgSectLbl.FontName Else .name = DFLT_LBL_PROPSPACED_FONT_NAME
                        If MsgSectLbl.FontSize <> 0 Then .Size = MsgSectLbl.FontSize Else .Size = DFLT_LBL_PROPSPACED_FONT_SIZE
                    End If
                    If MsgSectLbl.FontItalic Then .Italic = True
                    If MsgSectLbl.FontBold Then .Bold = True
                    If MsgSectLbl.FontUnderline Then .Underline = True
                End With
                If MsgSectLbl.FontColor <> 0 Then .ForeColor = MsgSectLbl.FontColor Else .ForeColor = rgbBlack
            End With
            SectTxtFrm.Top = SectLbl.Top + SectLbl.Height
            IsApplied(SectLbl) = True
        Else
            SectTxtFrm.Top = 0
        End If
        
        If MsgSectTxt.MonoSpaced Then
            SetupMsgSectMonoSpaced  ' returns the maximum width required for monospaced section
        Else ' proportional spaced
            SetupMsgSectPropSpaced
        End If
        SectTxtBox.SelStart = 0
        
    End If
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupMsgSectMonoSpaced(Optional ByVal msg_append As Boolean = False, _
                                   Optional ByVal msg_append_margin As String = vbNullString, _
                                   Optional ByVal msg_text As String = vbNullString)
' ------------------------------------------------------------------------------
' Setup the current monospaced message section. When a text is explicitely
' provided (msg_text) the section is setup with this one, else with the MsgText
' content. When an explicit text is provided the text either replaces the text,
' which the default or the text is appended when (msg_append = True).
' Note 1: All top and height adjustments - except the one for the text box
'         itself are finally done by SizeAndPosition services when all
'         elements had been set up.
' Note 2: The optional arguments (msg_append) and (msg_text) are used with the
'         Monitor service which ma replace or add the provided text
' ------------------------------------------------------------------------------
Const PROC = "SetupMsgSectMonoSpaced"
    
    On Error GoTo eh
    Dim MaxWidthAreaFrame   As Single
    Dim MaxWidthSectFrame   As Single
    Dim MaxWidthTextFrame   As Single
    
    MaxWidthAreaFrame = FormWidthMaxUsable - 4
    MaxWidthSectFrame = MaxWidthAreaFrame
    MaxWidthTextFrame = MaxWidthSectFrame
    
    '~~ Keep record of the controls which had been applied
    IsApplied(AreaMsg) = True
    IsApplied(SectFrm) = True
    IsApplied(SectTxtFrm) = True:    MonoSpaced(SectTxtFrm) = True
    IsApplied(SectTxtBox) = True:    MonoSpaced(SectTxtBox) = True:  MonoSpacedTbx(SectTxtBox) = True
    
    If msg_text <> vbNullString Then MsgSectTxt.Text = msg_text
  
    With SectTxtBox
        With .Font
            If MsgSectTxt.FontName <> vbNullString Then .name = MsgSectTxt.FontName Else .name = DFLT_LBL_MONOSPACED_FONT_NAME
            If MsgSectTxt.FontSize <> 0 Then .Size = MsgSectTxt.FontSize Else .Size = DFLT_LBL_MONOSPACED_FONT_SIZE
            If .Bold <> MsgSectTxt.FontBold Then .Bold = MsgSectTxt.FontBold
            If .Italic <> MsgSectTxt.FontItalic Then .Italic = MsgSectTxt.FontItalic
            If .Underline <> MsgSectTxt.FontUnderline Then .Underline = MsgSectTxt.FontUnderline
        End With
        If .ForeColor <> MsgSectTxt.FontColor And MsgSectTxt.FontColor <> 0 Then .ForeColor = MsgSectTxt.FontColor
    End With
    
    AutoSizeTextBox as_tbx:=SectTxtBox _
                  , as_text:=MsgSectTxt.Text _
                  , as_width_limit:=0 _
                  , as_append:=msg_append _
                  , as_append_margin:=msg_append_margin
    
    With SectTxtBox
        .SelStart = 0
        .Left = siHmarginFrames
        SectTxtFrm.Left = siHmarginFrames
        SectTxtFrm.Height = .Top + .Height
    End With ' SectTxtBox
        
    '~~ The width may expand or shrink depending on the change of the displayed text
    '~~ However, it cannot expand beyond the maximum width calculated for the text frame
    NewWidth(SectTxtFrm) = Min(MaxWidthTextFrame, SectTxtBox.Width)
    SectFrm.Width = Min(MaxWidthSectFrame, SectTxtFrm.Width)
    AreaMsg.Width = Min(MaxWidthSectFrame, SectFrm.Width)
    FormWidth = AreaMsg.Width
    AdjustedParentsWidthAndHeight SectTxtBox
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SetupMsgSectPropSpaced(Optional ByVal msg_append As Boolean = False, _
                                   Optional ByVal msg_append_marging As String = vbNullString, _
                                   Optional ByVal msg_text As String = vbNullString)
' ------------------------------------------------------------------------------
' Setup the current proportional spaced section. When a text is explicitely
' provided (msg_text) the section is setup with this one, else with the property
' MsgText content. When an explicit text is provided the text either replaces
' the text, which the default or the text is appended when (msg_appen = True).
' Note 1: When this proportional spaced section is setup the message width is
'         regarded final. However, top and height adjustments - except the one
'         for the text box itself are finally done by SizeAndPosition
'         services when all elements had been set up.
' Note 2: The optional arguments (msg_append) and (msg_text) are used with the
'         Monitor service which ma replace or add the provided text
' ------------------------------------------------------------------------------
    Const PROC = "SetupMsgSectPropSpaced"
    
    On Error GoTo eh
    
    IsApplied(AreaMsg) = True
    IsApplied(SectFrm) = True
    IsApplied(SectTxtFrm) = True
    IsApplied(SectTxtBox) = True

    '~~ For proportional spaced message sections the width is determined by the Message area's width
    AreaMsg.Width = Me.InsideWidth
    SectFrm.Width = AreaMsg.Width
    SectTxtFrm.Width = SectFrm.Width - 5
'    Debug.Print "SectTxtFrm.Width = " & SectTxtFrm.Width
    
    AreaBttns.Top = AreaMsg.Top + AreaMsg.Height + 20
    Me.Height = AreaBttns.Top + AreaBttns.Height + 20
    
    If msg_text <> vbNullString Then MsgSectTxt.Text = msg_text
    
    With SectTxtBox
        With .Font
            If MsgSectTxt.FontName <> vbNullString Then .name = MsgSectTxt.FontName Else .name = DFLT_LBL_PROPSPACED_FONT_NAME
            If MsgSectTxt.FontSize <> 0 Then .Size = MsgSectTxt.FontSize Else .Size = DFLT_LBL_PROPSPACED_FONT_SIZE
            If .Bold <> MsgSectTxt.FontBold Then .Bold = MsgSectTxt.FontBold
            If .Italic <> MsgSectTxt.FontItalic Then .Italic = MsgSectTxt.FontItalic
            If .Underline <> MsgSectTxt.FontUnderline Then .Underline = MsgSectTxt.FontUnderline
        End With
        If .ForeColor <> MsgSectTxt.FontColor And MsgSectTxt.FontColor <> 0 Then .ForeColor = MsgSectTxt.FontColor
    End With
    
    AutoSizeTextBox as_tbx:=SectTxtBox _
                  , as_width_limit:=SectTxtFrm.Width _
                  , as_text:=MsgSectTxt.Text _
                  , as_append:=msg_append _
                  , as_append_margin:=msg_append_marging
    
    With SectTxtBox
        .SelStart = 0
        .Left = HSPACE_LEFT
        TimedDoEvents    ' to properly h-align the text
    End With
    
    SectTxtFrm.Height = SectTxtBox.Top + SectTxtBox.Height
    SectFrm.Height = SectTxtFrm.Top + SectTxtFrm.Height
    AreaMsg.Height = ContentHeight(AreaMsg)
    AreaBttns.Top = AreaMsg.Top + AreaMsg.Height + 20
    Me.Height = AreaBttns.Top + AreaBttns.Height + 20
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns1()
' ------------------------------------------------------------------------------
' Unify all applied/visible button's size by assigning the maximum width and
' height provided with their setup, and adjust their resulting left position.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns1"
    
    On Error GoTo eh
    Dim cllButtonRows   As Collection:      Set cllButtonRows = DsgnBttnRows
    Dim siLeft          As Single
    Dim frRow           As MSForms.Frame    ' Frame for the buttons in a row
    Dim vButton         As Variant
    Dim lRow            As Long
    Dim lButton         As Long
    Dim HeightAreaBttns As Single
    
    For lRow = 1 To cllButtonRows.Count
        siLeft = HSPACE_LEFTRIGHT_BUTTONS
        Set frRow = cllButtonRows(lRow)
        If IsApplied(frRow) Then
            For Each vButton In DsgnRowBttns(lRow)
                If IsApplied(vButton) Then
                    lButton = lButton + 1
                    With vButton
                        .Left = siLeft
                        .Width = siMaxButtonWidth
                        .Height = siMaxButtonHeight
                        .Top = siVmarginFrames
                        siLeft = .Left + .Width + siHmarginButtons
                        If IsNumeric(vMsgButtonDefault) Then
                            If lButton = vMsgButtonDefault Then .Default = True
                        Else
                            If .Caption = vMsgButtonDefault Then .Default = True
                        End If
                    End With
                End If
                HeightAreaBttns = HeightAreaBttns + siMaxButtonHeight + siHmarginButtons
            Next vButton
        End If
        frRow.Width = frRow.Width + HSPACE_LEFTRIGHT_BUTTONS
    Next lRow
    Me.Height = AreaMsg.Top + AreaMsg.Height + HeightAreaBttns
        
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns2Rows()
' ------------------------------------------------------------------------------
' Adjust all applied/visible button rows height to the maximum buttons height
' and the row frames width to the number of the displayed buttons considering a
' certain margin between the buttons (siHmarginButtons) and a margin at the
' left and the right.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns2Rows"
    
    On Error GoTo eh
    Dim BttnsFrm      As MSForms.Frame:   Set BttnsFrm = DsgnBttnsFrm
    Dim BttnRows        As Collection:      Set BttnRows = DsgnBttnRows
    Dim BttnRowFrame    As MSForms.Frame
    Dim siTop           As Single
    Dim v               As Variant
    Dim lButtons        As Long
    Dim siHeight        As Single
    Dim BttnsFrmWidth As Single
    Dim dct             As Dictionary:      Set dct = AppliedBttnRows
    
    '~~ Adjust button row's width and height
    siHeight = AppliedButtonRowHeight
    siTop = siVmarginFrames
    For Each v In dct
        Set BttnRowFrame = v
        lButtons = dct(v)
        If IsApplied(BttnRowFrame) Then
            With BttnRowFrame
                .Top = siTop
                .Height = siHeight
                '~~ Provide some extra space for the button's design
                BttnsFrmWidth = CInt((siMaxButtonWidth * lButtons) _
                               + (siHmarginButtons * (lButtons - 1)) _
                               + (siHmarginFrames * 2)) - siHmarginButtons + 7
                .Width = BttnsFrmWidth + (HSPACE_LEFTRIGHT_BUTTONS * 2)
                siTop = .Top + .Height + siVmarginButtons
            End With
        End If
    Next v
    Set dct = Nothing

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns3Frame()
' ------------------------------------------------------------------------------
' Adjust the frame around all button row frames to the maximum width calculated
' by the adjustment of each of the rows frame.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns3Frame"
    
    On Error GoTo eh
    Dim v               As Variant
    Dim siWidth    As Single
    Dim siHeight   As Single
    
    If IsApplied(BttnsFrm) Then
        siWidth = ContentWidth(BttnsFrm)
        siHeight = ContentHeight(BttnsFrm)
        With BttnsFrm
            .Top = 0
            BttnsFrm.Height = siHeight
            BttnsFrm.Width = siWidth
            '~~ Center all button rows within the buttons frame
            For Each v In DsgnBttnRows
                If IsApplied(v) Then
                    FrameCenterHorizontal center_frame:=v, within_frame:=BttnsFrm
                End If
            Next v
        End With
    End If
    AreaBttns.Height = ContentHeight(BttnsFrm)

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub SizeAndPosition2Bttns4Area()
' ------------------------------------------------------------------------------
' Adjust the buttons area frame in accordance with the buttons frame.
' ------------------------------------------------------------------------------
    Const PROC = "SizeAndPosition2Bttns4Area"
    
    On Error GoTo eh
    Dim siHeight    As Single
    Dim siWidth     As Single
    
    siHeight = ContentHeight(AreaBttns)
    siWidth = ContentWidth(AreaBttns)
    NewWidth(AreaBttns) = Min(siWidth, (siMsgWidthMax - 10))
    
    If Not ScrollH_Applied(AreaBttns) Then
        AreaBttns.Width = BttnsFrm.Left + BttnsFrm.Width + ScrollV_Width(AreaBttns)
    End If
    
    If Not ScrollH_Applied(AreaBttns) Then
        If Not ScrollV_Applied(AreaBttns) Then
            AreaBttns.Height = BttnsFrm.Top + BttnsFrm.Height + ScrollH_Height(AreaBttns)
        End If
    End If
    
    FormWidth = AreaBttns.Width + ScrollV_Width(AreaBttns)
    FrameCenterHorizontal center_frame:=AreaBttns, left_margin:=10
    
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub TimedDoEvents()

    TimerBegin
    ' Unfortunately the 'way faster DoEvents' method below does not have the desired effect in this module
    ' If GetQueueStatus(QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT) Then DoEvents
    DoEvents ' this is way slower
'#If Debugging = 1 Then
''    Debug.Print "DoEvents in '" & tde_source & "' interrupted the code execution for " & TimerEnd & " msec"
'#End If

End Sub

Public Sub TimerBegin()
    cyTimerTicksBegin = TimerSysCurrentTicks
End Sub

Public Function TimerEnd() As Currency
    cyTimerTicksEnd = TimerSysCurrentTicks
    TimerEnd = TimerSecsElapsed * 1000
End Function

Private Sub UserForm_Activate()
' -------------------------------------------------------------------------------
' To avoid screen flicker the setup may has been done already. However for test
' purpose the Setup may run with the Activate event i.e. the .Show
' -------------------------------------------------------------------------------
    If Not bSetUpDone Then Setup
'    mMsg.MakeFormResizable
End Sub

Private Sub VisualizeControlsSetup()
' -------------------------------------------------------------------------------
'
' -------------------------------------------------------------------------------
    
    Dim ctl         As MSForms.Control
    Dim lBackColor  As Long
    Dim frm         As MSForms.Frame
    Dim tbx         As MSForms.TextBox
    Dim lbl         As MSForms.Label
    Dim indicate    As Boolean: indicate = bVisualizeControls
    
    lBackColor = Me.BackColor
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "Frame"
                Set frm = ctl
                With frm
                    If indicate Then
                        .BorderColor = -2147483638   ' active frame, allows with style none to hide the frame
                        .BorderStyle = fmBorderStyleSingle
                    Else
                        .BorderColor = lBackColor
                        .BorderStyle = fmBorderStyleNone
                        .BackColor = lBackColor
                    End If
                End With
            Case "TextBox"
                Set tbx = ctl
                With tbx
                    If indicate Then
                        .BorderColor = -2147483638   ' active frame, allows with style none to hide the frame
                        .BorderStyle = fmBorderStyleSingle
                    Else
                        .BorderColor = lBackColor
                        .BorderStyle = fmBorderStyleNone
                        .BackColor = lBackColor
                    End If
                End With
            Case "Label"
                Set lbl = ctl
                With lbl
                    If indicate Then
                        .BorderColor = -2147483638   ' active frame, allows with style none to hide the frame
                        .BorderStyle = fmBorderStyleSingle
                    Else
                        .BorderColor = lBackColor
                        .BorderStyle = fmBorderStyleNone
                        .BackColor = lBackColor
                    End If
                End With
        End Select
    Next ctl
    
End Sub

