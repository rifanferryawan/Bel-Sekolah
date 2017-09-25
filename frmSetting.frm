VERSION 5.00
Begin VB.Form frmSetting 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Jam Bel Sekolah"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin BelSekolah.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   4200
      TabIndex        =   77
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse2 
      Caption         =   "Bro&wse.."
      Height          =   375
      Left            =   5520
      TabIndex        =   76
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse1 
      Caption         =   "&Browse.."
      Height          =   375
      Left            =   5520
      TabIndex        =   75
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   5640
      Width           =   3615
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   5160
      Width           =   3615
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "&Keluar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   70
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   31
      Left            =   5280
      TabIndex        =   69
      Text            =   "11:45:00"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   30
      Left            =   5280
      TabIndex        =   68
      Text            =   "11:15:00"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   29
      Left            =   5280
      TabIndex        =   67
      Text            =   "10:45:00"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   28
      Left            =   5280
      TabIndex        =   63
      Text            =   "10:15:00"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   27
      Left            =   5280
      TabIndex        =   62
      Text            =   "09:45:00"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   26
      Left            =   5280
      TabIndex        =   61
      Text            =   "09:30:00"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   25
      Left            =   5280
      TabIndex        =   60
      Text            =   "09:00:00"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   24
      Left            =   5280
      TabIndex        =   59
      Text            =   "08:30:00"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   23
      Left            =   5280
      TabIndex        =   58
      Text            =   "08:00:00"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   22
      Left            =   5280
      TabIndex        =   57
      Text            =   "07:30:00"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   21
      Left            =   3240
      TabIndex        =   56
      Text            =   "13:25:00"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   20
      Left            =   3240
      TabIndex        =   55
      Text            =   "12:45:00"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   19
      Left            =   3240
      TabIndex        =   54
      Text            =   "11:55:00"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   18
      Left            =   3240
      TabIndex        =   53
      Text            =   "11:45:00"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   17
      Left            =   3240
      TabIndex        =   52
      Text            =   "11:15:00"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   16
      Left            =   3240
      TabIndex        =   51
      Text            =   "10:35:00"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   15
      Left            =   3240
      TabIndex        =   50
      Text            =   "09:55:00"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   14
      Left            =   3240
      TabIndex        =   49
      Text            =   "09:45:00"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   13
      Left            =   3240
      TabIndex        =   48
      Text            =   "09:05:00"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   12
      Left            =   3240
      TabIndex        =   47
      Text            =   "08:25:00"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   11
      Left            =   3240
      TabIndex        =   46
      Text            =   "07:45:00"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   10
      Left            =   1080
      TabIndex        =   45
      Text            =   "13:20:00"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   44
      Text            =   "12:35:00"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   43
      Text            =   "11:50:00"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   42
      Text            =   "11:40:00"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   41
      Text            =   "10:55:00"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   40
      Text            =   "10:10:00"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   39
      Text            =   "09:25:00"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   38
      Text            =   "09:15:00"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   37
      Text            =   "08:30:00"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   36
      Text            =   "07:45:00"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtJam 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   35
      Text            =   "07:00:00"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bel Masuk dan Pulang"
      Height          =   195
      Index           =   39
      Left            =   120
      TabIndex        =   74
      Top             =   5640
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bel Jam Pelajaran"
      Height          =   195
      Index           =   38
      Left            =   120
      TabIndex        =   73
      Top             =   5160
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   195
      Index           =   37
      Left            =   4560
      TabIndex        =   66
      Top             =   4200
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   195
      Index           =   36
      Left            =   2400
      TabIndex        =   65
      Top             =   4560
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   195
      Index           =   33
      Left            =   240
      TabIndex        =   64
      Top             =   4560
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   195
      Index           =   35
      Left            =   4560
      TabIndex        =   34
      Top             =   3840
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   195
      Index           =   34
      Left            =   4560
      TabIndex        =   33
      Top             =   3480
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   195
      Index           =   32
      Left            =   4560
      TabIndex        =   32
      Top             =   3120
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   195
      Index           =   31
      Left            =   4560
      TabIndex        =   31
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   195
      Index           =   30
      Left            =   4560
      TabIndex        =   30
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Istirahat"
      Height          =   195
      Index           =   29
      Left            =   4560
      TabIndex        =   29
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   195
      Index           =   28
      Left            =   4560
      TabIndex        =   28
      Top             =   1680
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   195
      Index           =   27
      Left            =   4560
      TabIndex        =   27
      Top             =   1320
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   195
      Index           =   26
      Left            =   4560
      TabIndex        =   26
      Top             =   960
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      Height          =   195
      Index           =   25
      Left            =   4560
      TabIndex        =   25
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari Ramadhan"
      Height          =   195
      Index           =   24
      Left            =   4560
      TabIndex        =   24
      Top             =   360
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   195
      Index           =   23
      Left            =   2400
      TabIndex        =   23
      Top             =   4200
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   195
      Index           =   22
      Left            =   2400
      TabIndex        =   22
      Top             =   3840
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Istirahat"
      Height          =   195
      Index           =   21
      Left            =   2400
      TabIndex        =   21
      Top             =   3480
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   195
      Index           =   20
      Left            =   2400
      TabIndex        =   20
      Top             =   3120
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   195
      Index           =   19
      Left            =   2400
      TabIndex        =   19
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   195
      Index           =   18
      Left            =   2400
      TabIndex        =   18
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Istirahat"
      Height          =   195
      Index           =   17
      Left            =   2400
      TabIndex        =   17
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   195
      Index           =   16
      Left            =   2400
      TabIndex        =   16
      Top             =   1680
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   195
      Index           =   15
      Left            =   2400
      TabIndex        =   15
      Top             =   1320
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   195
      Index           =   14
      Left            =   2400
      TabIndex        =   14
      Top             =   960
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      Height          =   195
      Index           =   13
      Left            =   2400
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari Upacara"
      Height          =   195
      Index           =   12
      Left            =   2400
      TabIndex        =   12
      Top             =   360
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Istirahat"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Istirahat"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari Biasa"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LWA_COLORKEY = &H1
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const BM_SETSTATE = &HF3
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub cmdBrowse1_Click()
With CommonDialog1
    .ShowOpen
    .Filter = "File Suara (*.wav)|*.wav"
    txtPath(0).Text = .FileName
End With

End Sub
Private Sub cmdBrowse2_Click()
With CommonDialog1
    .ShowOpen
    .Filter = "File Suara (*.wav)|*.wav"
    txtPath(1).Text = .FileName
End With

End Sub

Private Sub cmdKeluar_Click()
frmTampilan.Show
Unload Me
End Sub

Private Sub cmdSimpan_Click()
For i = 0 To 31
    MasukKeINI "Jam", Str(i), txtJam(i).Text, TempatINI
Next i
If txtPath(0).Text = "" And txtPath(1).Text = "" Then
    MsgBox "Maaf text tanda bel harus diisi dengan mencarinya file Wav", vbQuestion + vbOKOnly, "Cari Wav"
    Exit Sub
Else
    MasukKeINI "Bell", "Suara1", txtPath(0).Text, TempatINI
    MasukKeINI "Bell", "Suara2", txtPath(1).Text, TempatINI
End If
    
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
Dim Data As String
Dim Bel1 As String
Dim Bel2 As String
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
For i = 0 To 31
    txtJam(i).Text = BacaDariINI("Jam", Str(i), "", TempatINI)
Next i
    txtPath(0).Text = BacaDariINI("Bell", "Suara1", "", TempatINI)
    txtPath(1).Text = BacaDariINI("Bell", "Suara2", "", TempatINI)
End Sub

