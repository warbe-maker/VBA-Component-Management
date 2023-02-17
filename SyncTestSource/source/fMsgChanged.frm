VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsgChanged 
   ClientHeight    =   14805
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   12390
   OleObjectBlob   =   "fMsgChanged.frx":0000
End
Attribute VB_Name = "fMsgChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' -------------------------------------------------------------------------------
' UserForm fMsg Provides all means for a message with up to 7 separated text
'               sections, either proportional- or mono-spaced, with an optional
'               label, and up to 7 reply buttons.
'
' Design:       Since the implementation is merely design driven its setup is
'               essential. Design changes must adhere to the concept.
'
' Public Properties:
' - IndicateFrameCaptions Test option, indicated the frame names
' - MinButtonWidth
' - MsgTitle               The title displayed in the window handle bar
' - MinButtonWidth         Minimum button width in pt
' - MsgButtonDefault       The number of the default button
' - MsgBttns               Buttons to be displayed, Collection provided by the
'                          mMsg.Buttons service
' - MsgHeightMax           Percentage of screen height
' - MsgHeightMin           Percentage of screen height
' - MsgLabel               A section's label
' - MsgWidthMax            Percentage of screen width
' - MsgWidthMin            Defaults to 400 pt. the absolute minimum is 200 pt
' - Text                   A section's text or a monitor header, monitor footer
'                          or monitor step text
' - VisualizeForTest       Test option, visualizes the controls via a specific
'                          BackColor
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
' W. Rauschenberger Berlin, Mar 2022 (last revision)
' -------------------------------------------------------------------------------



