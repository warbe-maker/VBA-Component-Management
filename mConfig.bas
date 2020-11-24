Attribute VB_Name = "mConfig"
Option Explicit
Option Private Module

Public Function CompManAddinPath( _
       Optional sTitle As String = "Select a folder for the CompMan Addin") As String
    CompManAddinPath = mBasic.SelectFolder(sTitle)
End Function

Public Function CommonComponentsBasePath( _
       Optional ccbp_title As String = "Select a folder for the Common Component Workbooks") As String
    CommonComponentsBasePath = mBasic.SelectFolder(ccbp_title)
End Function
