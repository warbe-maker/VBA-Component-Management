Attribute VB_Name = "mConfig"
Option Explicit
Option Private Module

Public Function CommonAddinsPath(Optional sTitle As String = "Select a folder for the CompMan Addin") As String
    CommonAddinsPath = mBasic.SelectFolder(sTitle)
End Function

Public Function CommonBasePath(Optional sTitle As String = "Select a folder for the Common Component Workbooks") As String
    CommonBasePath = mBasic.SelectFolder(sTitle)
End Function
