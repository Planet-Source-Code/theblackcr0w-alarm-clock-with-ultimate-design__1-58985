VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Reveil 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TheBlackCrow's Reveil Version 1.0 Janv. 2005"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   3600
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6105
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00847161&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D0CFC6&
      Height          =   210
      Left            =   2115
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Choisissez un mp3 .."
      Top             =   2175
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00847161&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D0CFC6&
      Height          =   210
      Left            =   2115
      TabIndex        =   2
      Text            =   "00:00:00"
      Top             =   2430
      Width           =   1545
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6360
      Top             =   1515
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6360
      Top             =   1995
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6360
      Top             =   2475
   End
   Begin VB.Image Image7 
      Height          =   300
      Left            =   840
      MouseIcon       =   "Form1.frx":4450C
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":4465E
      Top             =   1230
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   4605
      MouseIcon       =   "Form1.frx":45C30
      MousePointer    =   99  'Custom
      Top             =   2160
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   4350
      MouseIcon       =   "Form1.frx":45D82
      MousePointer    =   99  'Custom
      Top             =   2415
      Width           =   600
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   3690
      MouseIcon       =   "Form1.frx":45ED4
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   4500
      MouseIcon       =   "Form1.frx":46026
      MousePointer    =   99  'Custom
      Top             =   390
      Width           =   555
   End
   Begin VB.Image Image2 
      Height          =   120
      Left            =   3555
      MouseIcon       =   "Form1.frx":46178
      MousePointer    =   99  'Custom
      Top             =   390
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   2715
      MouseIcon       =   "Form1.frx":462CA
      MousePointer    =   99  'Custom
      Top             =   375
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2085
      TabIndex        =   1
      Top             =   2670
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D0CFC6&
      Height          =   735
      Left            =   2625
      TabIndex        =   0
      Top             =   1065
      Width           =   2415
   End
End
Attribute VB_Name = "Reveil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declares "fil", "music" and "arq" as String
Dim fil As String, music As String, arq As String
'Declare the API Function for playing MP3 files
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Dim rev As String

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub



Private Sub Form_Load()
CommonDialog1.InitDir = App.Path
Text1.Text = Time
fil = "MP3|*.mp3"  'FILTER
rev = ""
'Me.Show
Dim t As Single
Dim rtn As Long
'  rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
'  rtn = rtn Or WS_EX_LAYERED
'  SetWindowLong hwnd, GWL_EXSTYLE, rtn
'  SetLayeredWindowAttributes hwnd, 0, 192, LWA_ALPHA
'SetLayeredWindowAttributes hwnd, &H0, 0, LWA_COLORKEY
If Me.Picture <> 0 Then
  Call SetAutoRgn(Me)
End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
mciSendString "close CURRFILE", 0&, 0, 0
End Sub

Private Sub Image1_Click()
About.Show

End Sub

Private Sub Image2_Click()
Reveil.WindowState = 1
End Sub

Private Sub Image3_Click()
mciSendString "close CURRFILE", 0&, 0, 0
Unload Me
End
End Sub

Private Sub Image4_Click()
If music = "" Then
MsgBox "Choisissez d'abord une musique pour la sonnerie", vbExclamation, "Erreur TBC Reveil"
Exit Sub
Else
rev = Text1.Text
Timer3.Enabled = True
End If
End Sub

Private Sub Image5_Click()
rev = ""
Timer3.Enabled = False
mciSendString "close CURRFILE", 0&, 0, 0


End Sub

Private Sub Image6_Click()
CommonDialog1.Filter = fil 'Filter ... only mp3 files allowed
CommonDialog1.ShowOpen 'Shows the Open Dialog for any file

arq = CommonDialog1.FileName 'Capturing the path and filename
music = GetShortName(arq) 'converting WINDOWS FORMAT NAME to DOS FORMAT... see the Module1
Text2.Text = music
End Sub

Private Sub Image7_Click()
mciSendString "close CURRFILE", 0&, 0, 0
rev = ""
Timer3.Enabled = False
Image7.Visible = False
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Time
End Sub

Private Sub Timer2_Timer()
If rev = "" Then
Label2.Caption = "Le reveil est OFF"
Else
Label2.Caption = "Le reveil est ON à l'heure " & rev
End If
End Sub

Private Sub Timer3_Timer()
If rev = Time Then
mciSendString "open " & music & " type MPEGVideo alias CURRFILE", 0&, 0, 0
    mciSendString "play CURRFILE repeat", 0&, 0, 0
Timer3.Enabled = False
Image7.Visible = True
End If
End Sub
