Attribute VB_Name = "modINI"
Public TempatINI As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function BacaDariINI(sSection As String, sKey As String, sDefault As String, sIniFile As String)
    Dim sBuffer As String, lRet As Long
    sBuffer = String$(255, 0)
    lRet = GetPrivateProfileString(sSection, sKey, "", sBuffer, Len(sBuffer), sIniFile)
    If lRet = 0 Then
        If sDefault <> "" Then MasukKeINI sSection, sKey, sDefault, sIniFile
        BacaDariINI = sDefault
    Else
        BacaDariINI = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    End If
End Function
Public Function MasukKeINI(sSection As String, sKey As String, sValue As String, sIniFile As String) As Boolean
    Dim lRet As Long
    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    MasukKeINI = (lRet)
End Function

Public Sub Settingan()
TempatINI = App.Path & "\Bel5ekol@h.CC@"
End Sub
