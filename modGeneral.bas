Attribute VB_Name = "modGeneral"
Option Explicit

Const HKEY_LOCAL_ROOT = &H80000000
Const HKEY_LOCAL_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = _
         KEY_QUERY_VALUE + KEY_SET_VALUE + _
         KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
         KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Public Const Tempat = HKEY_LOCAL_MACHINE
Public Const SubTempat = "Software\Sodikin\BelSekolah"
                     
'Tipe Reg Key ROOT ...
Const ERROR_SUCCESS = 0
Const REG_SZ = 1     ' Unicode nul terminated string
Const REG_DWORD = 4  ' 32-bit number

Private Declare Function RegOpenKeyEx Lib _
        "advapi32" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib _
        "advapi32" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal lpReserved As Long, ByRef lpType As Long, _
        ByVal lpData As String, ByRef lpcbData As Long) _
        As Long
Private Declare Function RegCreateKey Lib _
        "advapi32.dll" Alias "RegCreateKeyA" _
        (ByVal hKey As Long, ByVal lpSubKey As _
        String, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib _
        "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegSetValueEx Lib _
        "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal _
        lpValueName As String, ByVal _
        Reserved As Long, ByVal dwType _
        As Long, lpData As Any, ByVal _
        cbData As Long) As Long
        
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Function SimpanReg(hKey As Long, strPath As String, _
strValue As String, strData As String)
Dim KeyHand As Long
Dim r As Long
r = RegCreateKey(hKey, strPath, KeyHand)
r = RegSetValueEx(KeyHand, strValue, 0, _
REG_SZ, ByVal strData, Len(strData))
r = RegCloseKey(KeyHand)
End Function
Public Function BacaReg(hKey As Long, strPath As String, strValue As String, strData As String)
On Error GoTo Error
Dim Data As Long
Data = GetKeyValue(hKey, _
           strPath, strValue, strData)
Exit Function
Error:
  MsgBox "Tidak ada informasi Registry", _
         vbInformation, "NIHIL"
End Function
Public Function HapusReg(rClass As Long, Path As String, sKey As String)
Dim hKey As Long
Dim Data As Long
  Data = RegOpenKeyEx(rClass, Path, 0, KEY_ALL_ACCESS, hKey)
  Data = RegDeleteValue(hKey, sKey)
  RegCloseKey hKey
End Function
Public Function GetKeyValue(KeyRoot As Long, _
                            KeyName As String, _
                            SubKeyRef As String, _
                            ByRef KeyVal As String) _
                            As Boolean
    Dim i As Long           ' Counter untuk looping
    Dim rc As Long          ' Code pengembalian
    Dim hKey As Long        ' Penanganan membuka Registry Key
    Dim hDepth As Long      '
    Dim KeyValType As Long  ' Tipe Data sebuah Registry Key
    Dim tmpVal As String    ' Penyimpanan sementara nilai Registry Key
    Dim KeyValSize As Long  ' Ukuran variabel Registry Key
    '------------------------------------------------------------
    ' Buka RegKey di bawah KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    ' Buka Registry Key
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    ' Penanganan Error...
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    ' Alokasi Variable Space
    tmpVal = String$(1024, 0)
    ' Penanda Variable Size
    KeyValSize = 1024
    '------------------------------------------------------------
    ' Ambil Nilai Registry Key ...
    '------------------------------------------------------------
    ' Ambil/Buat nilai Key
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)
    ' Penanganan Errors
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    ' Win95 Adds Null Terminated String...
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
        ' Null ditemukan, Extract dari String
        tmpVal = Left(tmpVal, KeyValSize - 1)
    Else  ' WinNT tidak bernilai Null Terminate String...
        ' Null tidak ditemukan, Extract String saja
        tmpVal = Left(tmpVal, KeyValSize)
    End If
    '------------------------------------------------------------
    ' Memeriksa nilai tipe Key untuk konversi ...
    '------------------------------------------------------------
    Select Case KeyValType  ' Cari tipe data...
    Case REG_SZ             ' Tipe data string Registry Key
        KeyVal = tmpVal     ' Copy nilai String
    Case REG_DWORD          ' Tipe data Double Word Registry Key
        ' Konversikan setiap bit
        For i = Len(tmpVal) To 1 Step -1
            ' Bangun nilai Char. Dengan Char.
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
        Next
        ' Konversi Double Word ke String
        KeyVal = Format$("&h" + KeyVal)
    End Select
    GetKeyValue = True      ' Pengembalian sukses
    rc = RegCloseKey(hKey)  ' Tutup Registry Key
    Exit Function           ' Keluar dari fungsi
GetKeyError:                ' Bersihkan memori jika terjadi error...
    KeyVal = ""             ' Set Return Val ke string kosong
    GetKeyValue = False     ' Pengembalian gagal
    rc = RegCloseKey(hKey)  ' Tutup Registry Key
End Function








