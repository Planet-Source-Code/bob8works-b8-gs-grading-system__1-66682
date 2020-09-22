VERSION 5.00
Begin VB.Form frmPicPrev 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Image Preview - [ Press ESC to Close] "
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgPic 
      Height          =   525
      Left            =   150
      Top             =   150
      Width           =   1245
   End
End
Attribute VB_Name = "frmPicPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowForm(ByRef iPic As Image)

    On Error Resume Next
    Set imgPic.Picture = iPic
    imgPic.Refresh
    
    Me.Width = imgPic.Width + 350
    Me.Height = imgPic.Height + 600
    Err.Clear

    Me.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub
