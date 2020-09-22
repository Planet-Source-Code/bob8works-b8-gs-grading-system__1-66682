VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vincent J. Jamero     bob8works@yahoo.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   3930
      TabIndex        =   4
      Top             =   3810
      Width           =   2820
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developer/s:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   2940
      TabIndex        =   3
      Top             =   3810
      Width           =   795
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   2940
      Picture         =   "frmSplash.frx":0000
      Top             =   3420
      Width           =   330
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   2910
      Picture         =   "frmSplash.frx":0122
      Top             =   3180
      Width           =   2220
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   4425
      Left            =   15
      Top             =   15
      Width           =   7365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT © BOB8WORKS 2006"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2940
      TabIndex        =   2
      Top             =   3570
      Width           =   2460
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Richard Mewett, Fredrik Qvarfort and Vlad Vissoultchev"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   450
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credits:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   5640
      TabIndex        =   1
      Top             =   240
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   2310
      Left            =   510
      Picture         =   "frmSplash.frx":058D
      Top             =   1260
      Width           =   2340
   End
   Begin VB.Image Image5 
      Height          =   1200
      Left            =   0
      Picture         =   "frmSplash.frx":3C43
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1305
      Left            =   0
      Top             =   3120
      Width           =   7695
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   0
      Picture         =   "frmSplash.frx":3DC5
      Stretch         =   -1  'True
      Top             =   2700
      Width           =   15360
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Const HWND_NOTOPMOST = -2


Dim tl As Integer

Public Function ShowSplash()
    
    'show form
    SetWindowPos Me.hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS
    Me.Show
    
    DoEvents
    DoEvents
    DoEvents
    
    'continue loading...
    Call modMain.Main_AfterSD

End Function


Public Function ShowForm()
    
    'show form
    Me.Show
End Function

Public Sub UnloadSplash()
    Me.Enabled = False
    Timer1.Enabled = True
End Sub


Private Sub Form_Deactivate()
    UnloadSplash
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        UnloadSplash
    End If
End Sub










Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS
End Sub



Private Sub Trans(Level As Integer)
        Dim Msg As Long

        On Error Resume Next
        
        Msg = GetWindowLong(Me.hwnd, G)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong Me.hwnd, G, Msg
        SetLayeredWindowAttributes Me.hwnd, 0, Level, LWA_ALPHA
        
End Sub

Private Sub Form_Load()
    tl = 256
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    SetWindowPos frmLogin.hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, FLAGS
End Sub

Private Sub Timer1_Timer()
    
    Trans tl
    On Error Resume Next
    tl = tl - 50
    'Me.Width = Me.Width - 1475

    If tl < 25 Then
        Timer1.Enabled = False
        tl = 50
        
        Unload Me

    End If
End Sub


