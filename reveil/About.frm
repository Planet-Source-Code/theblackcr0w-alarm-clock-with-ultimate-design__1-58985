VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A propos"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "About.frx":000C
   ScaleHeight     =   3495
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   345
      Left            =   2985
      MouseIcon       =   "About.frx":2DEAE
      MousePointer    =   99  'Custom
      Top             =   30
      Width           =   360
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
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

Private Sub Image1_Click()
Unload Me
End Sub
