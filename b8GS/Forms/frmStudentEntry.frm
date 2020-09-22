VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStudentEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading..."
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStudentEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin b8Controls4.b8Pic b8Pic1 
      Height          =   2235
      Left            =   4140
      TabIndex        =   21
      Top             =   690
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   3942
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   345
      Left            =   2940
      TabIndex        =   20
      Top             =   2580
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdPic 
      Left            =   5340
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.jpg"
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   345
      Left            =   2940
      TabIndex        =   19
      Top             =   2160
      Width           =   1095
   End
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
      TabIndex        =   15
      Top             =   0
      Width           =   10305
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Entry"
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
         TabIndex        =   16
         Top             =   90
         Width           =   1980
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmStudentEntry.frx":000C
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.CheckBox chkActive 
      BackColor       =   &H00F5F5F5&
      Caption         =   "Active"
      Height          =   315
      Left            =   1380
      TabIndex        =   10
      Top             =   4320
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.TextBox txtLastName 
      Height          =   315
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   3
      Top             =   3900
      Width           =   5040
   End
   Begin VB.TextBox txtMiddleName 
      Height          =   315
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   2
      Top             =   3480
      Width           =   5040
   End
   Begin VB.TextBox txtFirstName 
      Height          =   315
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   1
      Top             =   3060
      Width           =   5040
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5190
      TabIndex        =   5
      Top             =   4920
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
      Left            =   3630
      TabIndex        =   4
      Top             =   4920
      Width           =   1395
   End
   Begin VB.TextBox txtStudentID 
      Height          =   315
      Left            =   1410
      MaxLength       =   14
      TabIndex        =   0
      Top             =   660
      Width           =   1980
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   0
      TabIndex        =   17
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
      TabIndex        =   18
      Top             =   4770
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   53
      BorderColor1    =   15592425
      BorderColor2    =   16777215
   End
   Begin VB.Label lblModifiedDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3300
      TabIndex        =   14
      Top             =   4500
      Width           =   90
   End
   Begin VB.Label lblCreationDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3300
      TabIndex        =   13
      Top             =   4320
      Width           =   90
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Modified:"
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
      Height          =   195
      Left            =   2700
      TabIndex        =   12
      Top             =   4500
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Created: "
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
      Height          =   195
      Left            =   2700
      TabIndex        =   11
      Top             =   4320
      Width           =   555
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   3975
      Width           =   810
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   3555
      Width           =   960
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   3135
      Width           =   825
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   720
      Width           =   840
   End
End
Attribute VB_Name = "frmStudentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mFormState As String

Dim CurStudent As tStudent

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

Public Function ShowEdit(sStudentID As Long) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    CurStudent.StudentID = sStudentID
    
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

Private Sub cmdRemove_Click()
    
    Dim FSO As New FileSystemObject
    Dim sOldPicPath As String
    
    If b8Pic1.IsPicLoaded = True Then
        If MsgBox("Are you sure to remove this Student's picture?", vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then
            
            sOldPicPath = App.Path & "/studpics/" & Val(txtStudentID.Text) & ".img"
            If FSO.FileExists(sOldPicPath) = True Then
                FSO.DeleteFile sOldPicPath
            End If
            
            b8Pic1.LoadPic sOldPicPath
            
        End If
    End If
    
    Set FSO = Nothing
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

    'check form fields
    
    'student id
    If GetTxtVal(txtStudentID.Text) < 1 Then
        MsgBox "Please enter Student ID", vbExclamation
        HLTxt txtStudentID
        Exit Function
    End If
    
    'check firstname
    If IsEmpty(txtFirstName.Text) Then
        MsgBox "Please enter First Name.", vbExclamation
        HLTxt txtFirstName
        Exit Function
    End If
    
    'check middlename
    If IsEmpty(txtMiddleName.Text) Then
        MsgBox "Please enter Middle Name.", vbExclamation
        HLTxt txtMiddleName
        Exit Function
    End If
    
    'check LastName
    If IsEmpty(txtLastName.Text) Then
        MsgBox "Please enter Last Name.", vbExclamation
        HLTxt txtLastName
        Exit Function
    End If
        
    'set new Student
    CurStudent.FirstName = txtFirstName.Text
    CurStudent.MiddleName = txtMiddleName.Text
    CurStudent.LastName = txtLastName.Text
    CurStudent.ModifiedDate = Now
    'curStudent.CreationDate = Now
    CurStudent.Active = IIf(chkActive.Value = vbChecked, True, False)

    'try
    'add new Student
    If EditStudent(CurStudent) = True Then
        
        'save picture
        Dim FSO As New FileSystemObject
        Dim sNewPicPath As String
        
        sNewPicPath = App.Path & "/studpics/" & CurStudent.StudentID & ".img"
        
        If b8Pic1.IsPicLoaded Then
            If b8Pic1.PicPath <> sNewPicPath Then
                If FSO.FileExists(sNewPicPath) = True Then
                    FSO.DeleteFile sNewPicPath
                End If
            
                FSO.CopyFile b8Pic1.PicPath, sNewPicPath, True
            End If
        End If
        
        Set FSO = Nothing
        
        MsgBox "Student entry was successfuly updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Student entry.", vbExclamation
        'set flag
        mShowEdit = False
        
    End If
    
End Function

Private Function SaveAdd()
    
    Dim newStudent As tStudent
    Dim oldStudent As tStudent


    'check form field
    
    'student id
    If GetTxtVal(txtStudentID.Text) < 1 Then
        MsgBox "Please enter Student ID", vbExclamation
        HLTxt txtStudentID
        Exit Function
    End If
    'check student id duplication
    If GetStudentByID(GetTxtVal(txtStudentID.Text), oldStudent) = True Then
        MsgBox "Invalid Student ID. This ID is already used by other student named '" & oldStudent.FirstName & " " & oldStudent.MiddleName & " " & oldStudent.LastName & "'", vbExclamation
        HLTxt txtStudentID
        Exit Function
    End If
    
    'check firstname
    If IsEmpty(txtFirstName.Text) Then
        MsgBox "Please enter First Name.", vbExclamation
        HLTxt txtFirstName
        Exit Function
    End If
    
    'check middlename
    If IsEmpty(txtMiddleName.Text) Then
        MsgBox "Please enter Middle Name.", vbExclamation
        HLTxt txtMiddleName
        Exit Function
    End If
    
    'check LastName
    If IsEmpty(txtLastName.Text) Then
        MsgBox "Please enter Last Name.", vbExclamation
        HLTxt txtLastName
        Exit Function
    End If
        
    'set new Student
    newStudent.StudentID = Val(txtStudentID.Text)
    newStudent.FirstName = txtFirstName.Text
    newStudent.MiddleName = txtMiddleName.Text
    newStudent.LastName = txtLastName.Text
    newStudent.ModifiedDate = Now
    newStudent.CreationDate = Now
    newStudent.Active = IIf(chkActive.Value = vbChecked, True, False)
    
    
    'try
    'add new Student
    If AddStudent(newStudent) = True Then
        
        'save picture
        Dim FSO As New FileSystemObject
        Dim sNewPicPath As String
        
        sNewPicPath = App.Path & "/studpics/" & newStudent.StudentID & ".img"
        
        If b8Pic1.IsPicLoaded Then
            If FSO.FileExists(sNewPicPath) = True Then
                FSO.DeleteFile sNewPicPath
            End If
            FSO.CopyFile b8Pic1.PicPath, sNewPicPath, True
        End If
        
        Set FSO = Nothing
        
        
        'set flag
        mShowAdd = True
        
        If MsgBox("Do you want to enroll this student?", vbQuestion + vbOKCancel) = vbOK Then
            
            'close form and return
            Unload Me
            
            frmEnrollmentEntry.ShowAdd newStudent.StudentID
            
        Else
            'close form and return
            Unload Me
        End If
                        
    Else
    
        MsgBox "Unable to save new Student entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    
End Function



Private Sub cmdSelect_Click()
        
    Dim sPath As String
    
    
    'check
    If IsNumeric(txtStudentID.Text) = False Then
        MsgBox "Please enter valid Student ID first.", vbExclamation
        HLTxt txtStudentID
        Exit Sub
    End If
    
    If Val(txtStudentID.Text) < 1 Then
        MsgBox "Please enter valid Student ID first.", vbExclamation
        HLTxt txtStudentID
        Exit Sub
    End If
    
        
    
    On Error GoTo errh:
    cdPic.Filename = ""
    cdPic.Filter = "JPEG Files|*.jpg|GIF Files|*.gif|All Files|*.*"
    cdPic.ShowOpen
    sPath = cdPic.Filename
    
    If Len(Trim(sPath)) < 1 Then
        Exit Sub
    End If
    
    b8Pic1.LoadPic sPath

    Exit Sub
errh:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
    
    If isON = True Then
        Exit Sub
    End If
    isON = True
    
    Select Case mFormState
        Case "add"
        
            'set caption
            Me.Caption = "Add Student"
            
            lblCreationDate.Caption = Now
            lblModifiedDate.Caption = ""

       
            
        Case "edit"

            'get info
            If GetStudentByID(CurStudent.StudentID, CurStudent) = False Then
                'show failed
                MsgBox "Student entry with Student ID : '" & CurStudent.StudentID & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set caption
            Me.Caption = "Edit Student"
            
            'set form ui info
            txtStudentID.Text = CurStudent.StudentID
            txtFirstName.Text = CurStudent.FirstName
            txtMiddleName.Text = CurStudent.MiddleName
            txtLastName.Text = CurStudent.LastName
            chkActive.Value = IIf(CurStudent.Active, vbChecked, vbUnchecked)
            
            lblCreationDate.Caption = CurStudent.CreationDate
            lblModifiedDate.Caption = CurStudent.ModifiedDate
            
            'load picture
            b8Pic1.LoadPic App.Path & "/studpics/" & CurStudent.StudentID & ".img"
            
            'disable some controls
            txtStudentID.Enabled = False
            
            'set first focused item
            txtFirstName.SetFocus
            
    End Select
    
End Sub


Private Sub Form_Load()
    isON = False
End Sub

