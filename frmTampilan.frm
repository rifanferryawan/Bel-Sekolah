VERSION 5.00
Begin VB.Form frmTampilan 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bel Sekolah"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   ControlBox      =   0   'False
   Icon            =   "frmTampilan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGantiPassword 
      Caption         =   "&Ganti Password"
      Height          =   375
      Left            =   2520
      TabIndex        =   35
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Isikan Password"
      Height          =   735
      Left            =   2520
      TabIndex        =   33
      Top             =   3240
      Width           =   2895
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   34
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Timer trmBel 
      Interval        =   1000
      Left            =   4320
      Top             =   2280
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   2520
      TabIndex        =   32
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetting 
      Caption         =   "&Setting"
      Height          =   375
      Left            =   2520
      TabIndex        =   31
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTampilan.frx":2CFA
      Left            =   3480
      List            =   "frmTampilan.frx":2D04
      TabIndex        =   30
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtKet 
      Height          =   285
      Index           =   2
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtKet 
      Height          =   285
      Index           =   1
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtKet 
      Height          =   285
      Index           =   0
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   10
      Text            =   "07:00:00"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Text            =   "07:45:00"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   8
      Text            =   "08:30:00"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Text            =   "09:15:00"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   6
      Text            =   "09:25:00"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Text            =   "10:10:00"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   4
      Text            =   "10:55:00"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   7
      Left            =   1200
      TabIndex        =   3
      Text            =   "11:40:00"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   8
      Left            =   1200
      TabIndex        =   2
      Text            =   "11:50:00"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   1
      Text            =   "12:35:00"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   10
      Left            =   1200
      TabIndex        =   0
      Text            =   "13:20:00"
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   29
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pukul"
      Height          =   195
      Index           =   14
      Left            =   2520
      TabIndex        =   25
      Top             =   1320
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   195
      Index           =   13
      Left            =   2520
      TabIndex        =   24
      Top             =   960
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari"
      Height          =   195
      Index           =   12
      Left            =   2520
      TabIndex        =   23
      Top             =   600
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari Biasa"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   21
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   19
      Top             =   1560
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Istirahat"
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   17
      Top             =   2280
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   16
      Top             =   2640
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   195
      Index           =   8
      Left            =   360
      TabIndex        =   15
      Top             =   3000
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Istirahat"
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   14
      Top             =   3360
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   195
      Index           =   10
      Left            =   360
      TabIndex        =   13
      Top             =   3720
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   195
      Index           =   11
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   195
      Index           =   33
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Width           =   90
   End
End
Attribute VB_Name = "frmTampilan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bel1 As String, Bel2 As String
Const LWA_COLORKEY = &H1
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const BM_SETSTATE = &HF3
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const SE_PRIVILEGE_ENABLED = &H2
Const TokenPrivileges = 3
Const TOKEN_ASSIGN_PRIMARY = &H1
Const TOKEN_DUPLICATE = &H2
Const TOKEN_IMPERSONATE = &H4
Const TOKEN_QUERY = &H8
Const TOKEN_QUERY_SOURCE = &H10
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_ADJUST_GROUPS = &H40
Const TOKEN_ADJUST_DEFAULT = &H80
Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Const ANYSIZE_ARRAY = 1

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const WM_CLOSE = &H10
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Type Luid
    lowpart As Long
    highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    'pLuid As Luid
    pLuid As LARGE_INTEGER
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Public Function InitiateShutdownMachine(ByVal Machine As String, Optional Force As Variant, Optional Restart As Variant, Optional AllowLocalShutdown As Variant, Optional Delay As Variant, Optional Message As Variant) As Boolean
    Dim hProc As Long
    Dim OldTokenStuff As TOKEN_PRIVILEGES
    Dim OldTokenStuffLen As Long
    Dim NewTokenStuff As TOKEN_PRIVILEGES
    Dim NewTokenStuffLen As Long
    Dim pSize As Long
    If IsMissing(Force) Then Force = False
    If IsMissing(Restart) Then Restart = True
    If IsMissing(AllowLocalShutdown) Then AllowLocalShutdown = False
    If IsMissing(Delay) Then Delay = 0
    If IsMissing(Message) Then Message = ""
    If InStr(Machine, "\\") = 1 Then
        Machine = Right(Machine, Len(Machine) - 2)
    End If
    If (LCase(GetMyMachineName) = LCase(Machine)) Then
        If AllowLocalShutdown = False Then Exit Function
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hProc) = 0 Then
            MsgBox "OpenProcessToken Error: " & GetLastError()
            Exit Function
        End If
        If LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, OldTokenStuff.Privileges(0).pLuid) = 0 Then
            MsgBox "LookupPrivilegeValue Error: " & GetLastError()
            Exit Function
        End If
        NewTokenStuff = OldTokenStuff
        NewTokenStuff.PrivilegeCount = 1
        NewTokenStuff.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        NewTokenStuffLen = Len(NewTokenStuff)
        pSize = Len(NewTokenStuff)
        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, NewTokenStuffLen, OldTokenStuff, OldTokenStuffLen) = 0 Then
            MsgBox "AdjustTokenPrivileges Error: " & GetLastError()
            Exit Function
        End If
        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
        NewTokenStuff.Privileges(0).Attributes = 0
        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, Len(NewTokenStuff), OldTokenStuff, Len(OldTokenStuff)) = 0 Then
            Exit Function
        End If
    Else
        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
    End If
    InitiateShutdownMachine = True
End Function
Function GetMyMachineName() As String
    Dim sLen As Long
    GetMyMachineName = Space(100)
    sLen = 100
    If GetComputerName(GetMyMachineName, sLen) Then
        GetMyMachineName = Left(GetMyMachineName, sLen)
    End If
End Function

Private Sub cmdGantiPassword_Click()
Dim Pass As String
Pass = BacaDariINI("Password", "Pass", "", TempatINI)
If txtPassword.Text = "" Then
    MsgBox "Maaf anda harus mengisi Password"
    txtPassword.SetFocus
ElseIf txtPassword.Text = DecryptText(Pass, "Sodikin") Then
    gnpass = InputBox("Silahkan anda masukkan password baru", "Input")
    MasukKeINI "Password", "Pass", EncryptText(gnpass, "Sodikin"), TempatINI
Else
    MsgBox "Maaf password yang anda isikan salah"
    txtPassword.SetFocus
End If

End Sub

Private Sub cmdKeluar_Click()
Dim Pass As String
Pass = BacaDariINI("Password", "Pass", "", TempatINI)
If txtPassword.Text = "" Then
    MsgBox "Maaf anda harus mengisi Password"
    txtPassword.SetFocus
ElseIf txtPassword.Text = DecryptText(Pass, "Sodikin") Then
    End
Else
    MsgBox "Maaf password yang anda isikan salah"
    txtPassword.SetFocus
End If
End Sub

Private Sub cmdSetting_Click()
frmSetting.Show
End Sub

Private Sub Combo1_Change()
If Combo1.Text = "Biasa" Then
    If txtKet(0).Text = "Senin" Then
        tanya = MsgBox("Apakah hari ini mau ada upacara", vbQuestion + vbYesNo, "Upacara")
        If tanya = vbYes Then
            For i = 0 To 10
                txtJam(i).Text = BacaDariINI("Jam", Str(i + 11), "", TempatINI)
            Next i
        Else
            For i = 0 To 10
                txtJam(i).Text = BacaDariINI("Jam", Str(i), "", TempatINI)
            Next i
        End If
        Label1(2).Caption = "1"
        Label1(3).Caption = "2"
        Label1(4).Caption = "3"
        Label1(5).Caption = "Istirahat"
        Label1(6).Caption = "4"
        Label1(7).Caption = "5"
        Label1(8).Caption = "6"
        Label1(9).Caption = "Istirahat"
        Label1(10).Caption = "7"
        Label1(11).Caption = "8"
        Label1(33).Caption = "9"
        Label1(33).Visible = True
        txtJam(10).Visible = True
        For i = 0 To 10
            txtJam(i).Text = BacaDariINI("Jam", Str(i), "", TempatINI)
        Next i
    ElseIf txtKet(0).Text = "Minggu" Then
        Label1(33).Visible = False
        txtJam(10).Visible = False
    End If
Else
    For i = 0 To 9
        txtJam(i).Text = BacaDariINI("Jam", Str(i + 22), "", TempatINI)
    Next i
        Label1(2).Caption = "1"
        Label1(3).Caption = "2"
        Label1(4).Caption = "3"
        Label1(5).Caption = "4"
        Label1(6).Caption = "Istirahat"
        Label1(7).Caption = "5"
        Label1(8).Caption = "6"
        Label1(9).Caption = "7"
        Label1(10).Caption = "8"
        Label1(11).Caption = "9"
        Label1(33).Visible = False
        txtJam(10).Visible = False
End If
    MasukKeINI "Bulan", "Bulan", Combo1.Text, TempatINI
    Label1(0).Caption = "Hari " + Combo1.Text

End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Biasa" Then
    If txtKet(0).Text = "Senin" Then
        tanya = MsgBox("Apakah hari ini mau ada upacara", vbQuestion + vbYesNo, "Upacara")
        If tanya = vbYes Then
            For i = 0 To 10
                txtJam(i).Text = BacaDariINI("Jam", Str(i + 11), "", TempatINI)
            Next i
        Else
            For i = 0 To 10
                txtJam(i).Text = BacaDariINI("Jam", Str(i), "", TempatINI)
            Next i
        End If
        Label1(2).Caption = "1"
        Label1(3).Caption = "2"
        Label1(4).Caption = "3"
        Label1(5).Caption = "Istirahat"
        Label1(6).Caption = "4"
        Label1(7).Caption = "5"
        Label1(8).Caption = "6"
        Label1(9).Caption = "Istirahat"
        Label1(10).Caption = "7"
        Label1(11).Caption = "8"
        Label1(33).Caption = "9"
        Label1(33).Visible = True
        txtJam(10).Visible = True
        For i = 0 To 10
            txtJam(i).Text = BacaDariINI("Jam", Str(i), "", TempatINI)
        Next i
    ElseIf txtKet(0).Text = "Minggu" Then
        Label1(33).Visible = False
        txtJam(10).Visible = False
    End If
Else
    For i = 0 To 9
        txtJam(i).Text = BacaDariINI("Jam", Str(i + 22), "", TempatINI)
    Next i
        Label1(2).Caption = "1"
        Label1(3).Caption = "2"
        Label1(4).Caption = "3"
        Label1(5).Caption = "4"
        Label1(6).Caption = "Istirahat"
        Label1(7).Caption = "5"
        Label1(8).Caption = "6"
        Label1(9).Caption = "7"
        Label1(10).Caption = "8"
        Label1(11).Caption = "9"
        Label1(33).Visible = False
        txtJam(10).Visible = False
End If
    MasukKeINI "Bulan", "Bulan", Combo1.Text, TempatINI
    Label1(0).Caption = "Hari " + Combo1.Text
End Sub

Private Sub Form_Initialize()
    InitCommonControls

End Sub

Private Sub Form_Load()
Dim CLR As Long
Dim Ret As Long
On Error Resume Next
Call Settingan
  CLR = &HFF0000
  Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
  Ret = Ret Or WS_EX_LAYERED
  SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
  SetLayeredWindowAttributes Me.hwnd, CLR, 0, LWA_COLORKEY
App.TaskVisible = False
Hari = Weekday(Now, vbSunday)
If Hari = 1 Then
txtKet(0).Text = "Minggu"
ElseIf Hari = 2 Then
txtKet(0).Text = "Senin"
ElseIf Hari = 3 Then
txtKet(0).Text = "Selasa"
ElseIf Hari = 4 Then
txtKet(0).Text = "Rabu"
ElseIf Hari = 5 Then
txtKet(0).Text = "Kamis"
ElseIf Hari = 6 Then
txtKet(0).Text = "Jum'at"
ElseIf Hari = 7 Then
txtKet(0).Text = "Sabtu"
End If
Combo1.Text = BacaDariINI("Bulan", "Bulan", "", TempatINI)
If Combo1.Text = "Biasa" Then
    If txtKet(0).Text = "Senin" Then
        tanya = MsgBox("Apakah hari ini mau ada upacara", vbQuestion + vbYesNo, "Upacara")
        If tanya = vbYes Then
            For i = 0 To 10
                txtJam(i).Text = BacaDariINI("Jam", Str(i + 11), "", TempatINI)
            Next i
        Else
            For i = 0 To 10
                txtJam(i).Text = BacaDariINI("Jam", Str(i), "", TempatINI)
            Next i
        End If
    End If
Else
    For i = 0 To 9
        txtJam(i).Text = BacaDariINI("Jam", Str(i + 22), "", TempatINI)
    Next i
End If
    
End Sub

Private Sub trmBel_Timer()
Bel1 = BacaDariINI("Bell", "Suara1", "", TempatINI)
Bel2 = BacaDariINI("Bell", "Suara2", "", TempatINI)
txtKet(1).Text = Format(Now, "dd mmmm yyyy")
txtKet(2).Text = Format(Now, "hh:mm:ss")
If Combo1.Text = "Biasa" Then
    If txtKet(0).Text = "Minggu" Then
        For i = 1 To 8
            If txtKet(2).Text = txtJam(i).Text Then
                sndPlaySound Bel1, 1
            End If
        Next i
        If txtKet(2).Text = txtJam(0).Text Or txtKet(2).Text = txtJam(9).Text Then
            sndPlaySound Bel2, 1
        End If
    Else
        For i = 1 To 9
            If txtKet(2).Text = txtJam(i).Text Then
                sndPlaySound Bel1, 1
            End If
        Next i
        If txtKet(2).Text = txtJam(0).Text Or txtKet(2).Text = txtJam(10).Text Then
            sndPlaySound Bel2, 1
        End If
    End If
Else
    For i = 1 To 8
        If txtKet(2).Text = txtJam(i).Text Then
            sndPlaySound Bel1, 1
        End If
    Next i
    If txtKet(2).Text = txtJam(0).Text Or txtKet(2).Text = txtJam(9).Text Then
        sndPlaySound Bel2, 1
    End If
End If
If txtKet(2).Text = "13:30:00" Then
InitiateShutdownMachine GetMyMachineName, True, False, True
End If
End Sub
Private Sub txtPassword_GotFocus()
cmdKeluar.Default = True
End Sub
