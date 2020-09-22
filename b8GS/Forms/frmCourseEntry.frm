VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmCourseEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading..."
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCourseEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
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
      TabIndex        =   8
      Top             =   0
      Width           =   10305
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Course Entry"
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
         TabIndex        =   9
         Top             =   90
         Width           =   1845
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmCourseEntry.frx":000C
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4170
      TabIndex        =   7
      Top             =   2280
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
      Left            =   2610
      TabIndex        =   6
      Top             =   2280
      Width           =   1395
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1620
      Width           =   3540
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1245
      Width           =   3540
   End
   Begin VB.TextBox txtCourseID 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1725
      TabIndex        =   0
      Top             =   840
      Width           =   1650
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   0
      TabIndex        =   10
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
      TabIndex        =   11
      Top             =   2130
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   53
      BorderColor1    =   15592425
      BorderColor2    =   16777215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   5
      Top             =   1710
      Width           =   855
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
      TabIndex        =   3
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Course ID:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   870
      Width           =   780
   End
End
Attribute VB_Name = "frmCourseEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mFormState As String

Dim curCourse As tCourse

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

Public Function ShowEdit(sCourseID As Integer) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curCourse.CourseID = sCourseID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function

Private Sub b83DRect1_DragDrop(Source As Control, x As Single, y As Single)

End Sub

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

    Dim newCourse As tCourse
    Dim oldCourse As tCourse
    
    Dim iCourseID As Integer
    Dim iPackageID As Integer
    
    'check form field
    If IsEmpty(txtCourseID.Text) Then
        MsgBox "Please enter Course ID", vbExclamation
        txtCourseID.Text = GetNewCourseID
        Exit Function
    End If
    
    If IsEmpty(txtTitle.Text) Then
        MsgBox "Please enter Name", vbExclamation
        HLTxt txtCourseID
        Exit Function
    End If
    
    'set new Course
    'curCourse.CourseID = GetTxtVal(txtCourseID.Text)
    curCourse.Title = txtTitle.Text
    curCourse.Description = txtDescription.Text
    
    'try
    'add new Course
    If EditCourse(curCourse) = True Then
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Course entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function

Private Function SaveAdd()
    Dim newCourse As tCourse
    Dim oldCourse As tCourse
    
    Dim iCourseID As Integer
    Dim iPackageID As Integer
    
    'check form field
    If IsEmpty(txtCourseID.Text) Then
        MsgBox "Please enter Course ID", vbExclamation
        txtCourseID.Text = GetNewCourseID
        Exit Function
    End If
    If IsEmpty(txtTitle.Text) Then
        MsgBox "Please enter Name", vbExclamation
        HLTxt txtTitle
        Exit Function
    End If

    'check Course
    On Error Resume Next
    
    'check title duplication
    If GetCourseByTitle(txtTitle.Text, oldCourse) = True Then
        MsgBox "The Title that you have entered was already exist in record." & vbNewLine & _
            "Please enter another Title.", vbExclamation
            
        HLTxt txtTitle
        Exit Function
    End If
    
    'set new Course
    newCourse.CourseID = Val(txtCourseID.Text)
    newCourse.Title = txtTitle.Text
    newCourse.Description = txtDescription.Text
    
    
    'try
    'add new Course
    If AddCourse(newCourse) = True Then
        MsgBox "New Course entry was added.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to add new Course entry.", vbExclamation
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
            Me.Caption = "Add Course"
            
            'generate ID
            txtCourseID.Text = GetNewCourseID

           
            
        Case "edit"

            'get info
            If GetCourseByID(curCourse.CourseID, curCourse) = False Then
                'show failed
                MsgBox "Course entry with Course ID : '" & curCourse.CourseID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            txtCourseID.Text = curCourse.CourseID
            txtTitle.Text = curCourse.Title
            txtDescription.Text = curCourse.Description

            'set caption
            Me.Caption = "Edit Course"

            
    End Select
    
End Sub


Private Sub Form_Load()
    isON = False
End Sub

