VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "User's Login"
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin b8Controls4.b8GradLine b8GradLine1 
      Height          =   30
      Left            =   3660
      TabIndex        =   8
      Top             =   1500
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   53
      Color1          =   12632256
      Color2          =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imglUser 
      Left            =   3780
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":058A
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":0B24
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo txtUserID 
      Height          =   375
      Left            =   3690
      TabIndex        =   1
      Top             =   1830
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      ImageList       =   "imglUser"
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
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
      Left            =   4230
      TabIndex        =   4
      Top             =   3870
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5790
      TabIndex        =   5
      Top             =   3870
      Width           =   1395
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3690
      MaxLength       =   20
      PasswordChar    =   "="
      TabIndex        =   3
      Top             =   3210
      Width           =   3285
   End
   Begin b8Controls4.b8GradLine b8GradLine2 
      Height          =   30
      Left            =   3660
      TabIndex        =   9
      Top             =   2880
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   53
      Color1          =   12632256
      Color2          =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   2280
      Left            =   60
      Picture         =   "frmLogin.frx":10BE
      Top             =   2070
      Width           =   2460
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User's Log-in"
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
      Left            =   120
      TabIndex        =   12
      Top             =   690
      Width           =   1875
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   3660
      TabIndex        =   11
      Top             =   2670
      Width           =   1125
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[2]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   3180
      TabIndex        =   10
      Top             =   2580
      Width           =   450
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   3690
      TabIndex        =   2
      Top             =   2970
      Width           =   855
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[1]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Left            =   3180
      TabIndex        =   7
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Account"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   3660
      TabIndex        =   6
      Top             =   1290
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":36A4
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   660
      Picture         =   "frmLogin.frx":539E
      Top             =   450
      Width           =   330
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   630
      Picture         =   "frmLogin.frx":54C0
      Top             =   270
      Width           =   1545
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&User ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   3690
      TabIndex        =   0
      Top             =   1590
      Width           =   675
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   4425
      Left            =   15
      Top             =   15
      Width           =   7365
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mShowForm As Boolean
Dim dFailedCount As Integer

Public Function ShowForm() As Boolean

    'load user accounts
    RefreshUsers GetSetting(App.EXEName, "TextBox", txtUserID.Name, "")
    
    'show form
    Me.Show vbModal
    
    ShowForm = mShowForm
End Function

Private Sub cmdCancel_Click()
    mShowForm = False
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    
    'check form field
    If IsEmpty(txtUserID.Text) Then
        MsgBox "Please enter User ID", vbExclamation
        HLTxt txtUserID
        Exit Sub
    End If
    
    If IsEmpty(txtPassword.Text) Then
        MsgBox "Please enter Password", vbExclamation
        HLTxt txtPassword
        Exit Sub
    End If
    
    'check user
    If GetUserByID(txtUserID.Text, CurrentUser) = False Then
        MsgBox "User does not exist.", vbExclamation
        HLTxt txtUserID
        Exit Sub
    End If
    
    'check password
    If txtPassword.Text <> CurrentUser.Password Then
    
        If dFailedCount >= 5 Then
            WriteErrorLog Me.Name, "cmdLogin_Click", "Err: 0000x000FF"
            Unload Me
            Exit Sub
        End If
    
        MsgBox "Invalid Password.", vbExclamation
        HLTxt txtPassword
        
        'increment counter
        dFailedCount = dFailedCount + 1
        
        Exit Sub
    End If
    
    
    'set current user
    If GetUserByID(Trim(txtUserID.Text), CurrentUser) = False Then
        WriteErrorLog Me.Name, "cmdLogin_Click", "GetUserByID(Trim(txtUserID.Text), CurrentUser) = False"
        Unload Me
    End If
    
    
    'success
    'write to log
    'temp
    
    'set flag
    mShowForm = True
    'close this form
    Unload Me
    
End Sub


Private Sub Form_Load()
    'default
    dFailedCount = 0
    
     
    PaintGrad Me, &HE0E0E0, &HFFFFFF, 135
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    modFunction.FormDrag Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting App.EXEName, "TextBox", txtUserID.Name, txtUserID.Text
End Sub





Private Sub RefreshUsers(Optional sUserID As String = "")

    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer

    txtUserID.ComboItems.Clear
    
    sSQL = "SELECT tblUser.UserID, tblUser.Password, tblUser.CreationDate, tblUser.CreatedBy, tblUser.ModifiedDate, tblUser.ModifiedBy" & _
            " FROM tblUser"

    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "RefreshUsers", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo ReleaseAndExit
    End If
    
    If AnyRecordExisted(vRS) = False Then
        MsgBox "The are no User yet. Please add some User first.", vbExclamation
        GoTo ReleaseAndExit
    End If

    i = 1
    vRS.MoveFirst
    While vRS.EOF = False
        
        txtUserID.ComboItems.Add i, , ReadField(vRS.Fields("UserID")), IIf(LCase(Trim(ReadField(vRS.Fields("UserID")))) = "administrator", "admin", "user")
        If LCase(sUserID) = LCase(txtUserID.ComboItems(i).Text) Then
            txtUserID.ComboItems(i).Selected = True
        End If
        vRS.MoveNext
        i = i + 1
    Wend
    
    
    
ReleaseAndExit:
    Set vRS = Nothing
End Sub


