Attribute VB_Name = "mRegTest"
Option Explicit

Public Sub Test_Reg()
    Dim s As String
    
    '~~ Test 1: Entry exists (key or name)
    Debug.Assert mReg.Exists("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot\") = True

    '~~ Test 2: Write a value
    Debug.Assert mReg.Exists("HKCU\CompMan\BasicConfig\ServicedRootFolder") = False
    mReg.Value("HKCU\CompMan\BasicConfig\ServicedRootFolder") = "E:\Ablage\Excel VBA\DevAndTest"
    Debug.Assert mReg.Exists("HKCU\CompMan\BasicConfig\ServicedRootFolder") = True
    
    '~~ Test 3: Read a value
    s = mReg.Value("HKCU\CompMan\BasicConfig\ServicedRootFolder")
    Debug.Assert s = "E:\Ablage\Excel VBA\DevAndTest"

    '~~ Test 4: Delete key
    Debug.Assert mReg.Exists("HKCU\CompMan\BasicConfig\ServicedRootFolder") = True
    mReg.Delete "HKCU\CompMan\BasicConfig\ServicedRootFolder"
    Debug.Assert mReg.Exists("HKCU\CompMan\BasicConfig\ServicedRootFolder") = False

End Sub

