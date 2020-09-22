VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmSYEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading..."
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSYEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   158
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   6
      Top             =   0
      Width           =   10305
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S.Y. Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00926747&
         Height          =   345
         Left            =   570
         TabIndex        =   7
         Top             =   90
         Width           =   1410
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmSYEntry.frx":000C
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2850
      TabIndex        =   5
      Top             =   1860
      Width           =   1395
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1290
      TabIndex        =   4
      Top             =   1860
      Width           =   1395
   End
   Begin VB.TextBox txtSYTitle 
      Height          =   315
      Left            =   1725
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1215
      Width           =   2340
   End
   Begin VB.TextBox txtSYID 
      Height          =   315
      Left            =   1725
      TabIndex        =   1
      Top             =   765
      Width           =   2340
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   555
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   53
      BorderColor1    =   15592425
      BorderColor2    =   16777215
   End
   Begin b8Controls4.b8Line b8Line2 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   1710
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   53
      BorderColor1    =   15592425
      BorderColor2    =   16777215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S.Y. ID."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   840
      Width           =   570
   End
End
Attribute VB_Name = "frmSYEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mFormState As String

Dim curSY As tSY

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean
Dim isON As Boolean

Public Function ShowAdd() As Boolean
    
    'set form state
    mFormState = "add"
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowEdit(sSYID As Integer) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curSY.SYID = sSYID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function

Private Sub cmdCancel_Click()
    
    Select Case mFormState
        Case "add"
            mShowAdd = False
        Case "edit"
            mShowEdit = False
    End Select
    
    Unload Me
End Sub

Private Sub cmdSave_Click()

    Select Case mFormState
        Case "add"
            SaveAdd
        Case "edit"
            SaveEdit
    End Select
    
End Sub

Private Function SaveEdit()

    Dim newSY As tSY
    Dim oldSY As tSY
    
    Dim iSYID As Integer
    Dim iPackageID As Integer
    
    'check form field
    If IsEmpty(txtSYID.Text) Then
        MsgBox "Please enter SY ID", vbExclamation
        HLTxt txtSYID
        Exit Function
    End If
    
    If IsEmpty(txtSYTitle.Text) Then
        MsgBox "Please enter Name", vbExclamation
        HLTxt txtSYID
        Exit Function
    End If
    
    'set new SY
    'curSY.SYID = GetTxtVal(txtSYID.Text)
    curSY.SYTitle = txtSYTitle.Text
    
    'try
    'add new SY
    If EditSY(curSY) = True Then
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update SY entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function

Private Function SaveAdd()
    Dim newSY As tSY
    Dim oldSY As tSY
    
    Dim iSYID As Integer
    Dim iPackageID As Integer
    
    'check form field
    If IsEmpty(txtSYID.Text) Then
        MsgBox "Please enter SY ID", vbExclamation
        HLTxt txtSYID
        Exit Function
    End If
    If IsEmpty(txtSYTitle.Text) Then
        MsgBox "Please enter Name", vbExclamation
        HLTxt txtSYTitle
        Exit Function
    End If

    'check SY
    On Error Resume Next
    
    'check title duplication
    If GetSYBySYTitle(txtSYTitle.Text, oldSY) = True Then
        MsgBox "The Title that you have entered was already exist in record." & vbNewLine & _
            "Please enter another Title.", vbExclamation
            
        HLTxt txtSYID
        Exit Function
    End If
    
    'set new SY
    newSY.SYID = Val(txtSYID.Text)
    newSY.SYTitle = txtSYTitle.Text
    
    
    'try
    'add new SY
    If AddSY(newSY) = True Then
        MsgBox "New SY entry was added.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to add new SY entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function



Private Sub Form_Activate()
    
    If isON = True Then
        Exit Sub
    End If
    isON = True
    
    Select Case mFormState
        Case "add"
        
            'set caption
            Me.Caption = "Add SY"

           
            
        Case "edit"

            'get info
            If GetSYByID(curSY.SYID, curSY) = False Then
                'show failed
                MsgBox "SY entry with SY ID : '" & curSY.SYID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            txtSYID.Text = curSY.SYID
            txtSYTitle.Text = curSY.SYTitle

            'set caption
            Me.Caption = "Edit SY"

            
    End Select
    
End Sub


Private Sub Form_Load()
    isON = False
End Sub

Private Sub txtSYID_Change()
    
    txtSYTitle.Text = ""
    
    If IsNumeric(txtSYID.Text) = False Then
        Exit Sub
    End If
    
    If Val(txtSYID.Text) < 1000 Or Val(txtSYID.Text) > 9998 Then
        Exit Sub
    End If
    
    txtSYTitle.Text = GetSYTitle(Val(txtSYID.Text))
    
End Sub

