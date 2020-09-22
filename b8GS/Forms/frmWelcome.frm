VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Welcome"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image4 
      Height          =   135
      Left            =   3660
      Picture         =   "frmWelcome.frx":0000
      Top             =   4650
      Width           =   330
   End
   Begin VB.Image Image5 
      Height          =   1200
      Left            =   0
      Picture         =   "frmWelcome.frx":0122
      Stretch         =   -1  'True
      Top             =   60
      Width           =   15360
   End
   Begin VB.Image Image1 
      Height          =   2310
      Left            =   990
      Picture         =   "frmWelcome.frx":02A4
      Top             =   2490
      Width           =   2340
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   3630
      Picture         =   "frmWelcome.frx":395A
      Top             =   4410
      Width           =   2220
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   0
      Picture         =   "frmWelcome.frx":3DC5
      Stretch         =   -1  'True
      Top             =   3930
      Width           =   15360
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   12000
      Left            =   0
      Top             =   4350
      Width           =   15360
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowForm()


    mdiMain.AddChild Me, False
    
End Sub

Private Sub Form_Activate()
    mdiMain.ActivateChild Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'pass keyinfo to mdiMain
    mdiMain.AFForm_KeyDown KeyCode, Shift
End Sub
