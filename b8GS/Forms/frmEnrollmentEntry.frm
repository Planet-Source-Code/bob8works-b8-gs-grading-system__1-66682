VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmEnrollmentEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading..."
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEnrollmentEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
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
      TabIndex        =   17
      Top             =   0
      Width           =   10305
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enrollment Entry"
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
         TabIndex        =   18
         Top             =   90
         Width           =   2430
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmEnrollmentEntry.frx":000C
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.CheckBox chkIncludeNonActive 
      BackColor       =   &H00F5F5F5&
      Caption         =   "&Include Non-Active"
      Height          =   225
      Left            =   1260
      TabIndex        =   16
      Top             =   1740
      Width           =   1695
   End
   Begin b8Controls4.b8DataPicker b8DPStudent 
      Height          =   345
      Left            =   1260
      TabIndex        =   15
      Top             =   1380
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   609
      SQLWhereSeparator=   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtEnrollmentID 
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   840
      Width           =   2340
   End
   Begin VB.ComboBox cmbYL 
      Height          =   315
      Left            =   1260
      TabIndex        =   11
      Top             =   2220
      Width           =   1155
   End
   Begin VB.TextBox txtRemarks 
      Height          =   975
      Left            =   3720
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2820
      Width           =   3360
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5730
      TabIndex        =   1
      Top             =   4140
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
      Left            =   4140
      TabIndex        =   0
      Top             =   4140
      Width           =   1395
   End
   Begin MSComctlLib.ImageCombo cmbSY 
      Height          =   330
      Left            =   1260
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageCombo cmbCourse 
      Height          =   330
      Left            =   1260
      TabIndex        =   4
      Top             =   3060
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageCombo cmbSem 
      Height          =   330
      Left            =   1260
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   0
      TabIndex        =   19
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
      TabIndex        =   20
      Top             =   4020
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   53
      BorderColor1    =   15592425
      BorderColor2    =   16777215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enrollment ID:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   900
      Width           =   1020
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Y.L.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3720
      TabIndex        =   10
      Top             =   2580
      Width           =   675
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   1470
      Width           =   570
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sem.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   3540
      Width           =   420
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Course:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   3120
      Width           =   570
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S.Y.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   2700
      Width           =   360
   End
End
Attribute VB_Name = "frmEnrollmentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mFormState As String

Dim curEnrollment As tEnrollment

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isON As Boolean

Public Function ShowAdd(Optional lStudentID As Long = 0, Optional iSYID As Integer = 0, Optional iCourseID As Integer = 0) As Boolean
    
    'set form state
    mFormState = "add"
    
    'set parameter
    curEnrollment.FK_StudentID = lStudentID
    If iSYID > 0 Then
        curEnrollment.FK_SYID = iSYID
    Else
        curEnrollment.FK_SYID = modGV.CurrentSY.SYID
    End If
    curEnrollment.FK_CourseID = iCourseID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowEdit(ByVal sEnrollmentID As String) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curEnrollment.EnrollmentID = sEnrollmentID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function

Private Sub b8DPStudent_BeforeDropDown()
    If chkIncludeNonActive.Value = vbChecked Then
        b8DPStudent.SQLWhere = ""
    Else
        b8DPStudent.SQLWhere = " tblStudent.Active = True "
    End If
End Sub

Private Sub b8DPStudent_Change()
    GenerateEnrollmentID
End Sub

Private Sub cmbSem_Change()
    GenerateEnrollmentID
End Sub

Private Sub cmbSem_Click()
    GenerateEnrollmentID
End Sub

Private Sub cmbSY_Change()
    GenerateEnrollmentID
End Sub

Private Sub cmbSY_Click()
    GenerateEnrollmentID
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

Private Sub GenerateEnrollmentID()
    
    Dim tmpStudent As tStudent
    
    txtEnrollmentID.Text = ""
    
    If IsComboItemSelected(cmbSY) = False Then
        Exit Sub
    End If
    If IsComboItemSelected(cmbSem) = False Then
        Exit Sub
    End If
    
    'check student id
    If IsNumeric(b8DPStudent.BoundData) = True Then
        If Val(b8DPStudent.BoundData) < 2147483647 Then
            If GetStudentByID(CLng(b8DPStudent.BoundData), tmpStudent) = False Then
                Exit Sub
            End If
        End If
    End If
    
    
    'Enrollment ID = Student ID + "-" + SYID + "-" + Sem
    txtEnrollmentID.Text = Trim(GetTxtVal(b8DPStudent.BoundData)) & _
                            "-" & GetKeyOnSplit(cmbSY.SelectedItem.Key, "//", 1) & _
                            "-" & Trim(cmbSem.Text)
End Sub

Private Function SaveEdit()

    
End Function

Private Function SaveAdd()

    Dim NewEnrollment As tEnrollment
    Dim OldEnrollment As tEnrollment
    

    Dim vStudent As tStudent
    
    On Error GoTo RAE
    'checking
    
    'check student id
    If IsNumeric(b8DPStudent.BoundData) = False Then
        MsgBox "Please enter valid Student ID.", vbExclamation
        b8DPStudent.FocusedDropButton
        Exit Function
    End If
    If Val(b8DPStudent.BoundData) >= 2147483647 Then
        MsgBox "Please enter valid Student ID.", vbExclamation
        b8DPStudent.FocusedDropButton
        Exit Function
    End If
    If GetStudentByID(CLng(b8DPStudent.BoundData), vStudent) = False Then
        MsgBox "Student ID not found.", vbExclamation
        b8DPStudent.FocusedDropButton
        Exit Function
    End If

    'check sy
    If IsComboItemSelected(cmbSY) = False Then
        MsgBox "Please select valid S.Y.", vbExclamation
        cmbSY.SetFocus
        Exit Function
    End If
    
    'check course
    If IsComboItemSelected(cmbCourse) = False Then
        MsgBox "Please select valid Course.", vbExclamation
        cmbCourse.SetFocus
        Exit Function
    End If
    
    'check sem
    If IsComboItemSelected(cmbSem) = False Then
        MsgBox "Please select valid Sem", vbExclamation
        cmbSem.SetFocus
        Exit Function
    End If

    With NewEnrollment
        .EnrollmentID = txtEnrollmentID.Text
        .CreationDate = Now
        .FK_StudentID = CLng(b8DPStudent.BoundData)
        .FK_SYID = Val(GetKeyOnSplit(cmbSY.SelectedItem.Key, "//", 1))
        .FK_CourseID = Val(GetKeyOnSplit(cmbCourse.SelectedItem.Key, "//", 1))
        .YL = modYL.YLStrToNum(cmbYL.Text)
        .Sem = Val(cmbSem.Text)
        .ModifiedDate = Now
        .Remarks = txtRemarks.Text
    End With
    
    'check if student is allready enrolled
    If GetEnrollmentByID(NewEnrollment.EnrollmentID, OldEnrollment) = True Then
        MsgBox "The selected student is already enrolled.", vbExclamation
        b8DPStudent.FocusedDropButton
        Exit Function
    End If
    
    
    
    
    'Exit Function
    
    'try
    'add new Enrollment
    If AddEnrollment(NewEnrollment) = True Then
        
        'update yr level
        UpdateStudentYL NewEnrollment.FK_StudentID, NewEnrollment.YL
        
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to add new Enrollment entry.", vbExclamation
        'set flag
        mShowAdd = False
        
    End If
    Exit Function
    
RAE:
    WriteErrorLog Me.Name, "SaveAdd", Err.Description
End Function



Private Sub Form_Activate()
    
    Dim vStudent As tStudent
    
    If isON = True Then
        Exit Sub
    End If
    isON = True
    
    Select Case mFormState
        Case "add"
        
            'set caption
            Me.Caption = "Add Enrollment"
            
            If curEnrollment.FK_StudentID > 0 Then
                If GetStudentByID(curEnrollment.FK_StudentID, vStudent) = True Then
                b8DPStudent.BoundData = vStudent.StudentID
                b8DPStudent.DisplayData = vStudent.FirstName & " " & vStudent.MiddleName & " " & vStudent.LastName
                End If
            End If
            
            'refresh SY
            Refresh_SY curEnrollment.FK_SYID
            'refresh Courses
            Refresh_Course
            'refresh sem
            Refresh_Sem
            'refresh yl
            Refresh_YL

         
        Case "edit"
            
            'set caption
            Me.Caption = "Edit Enrollment"

            'get info
            'If GetEnrollmentByID(curEnrollment.FK_CourseID, curEnrollment.FK_SYID, curEnrollment.Sem, curEnrollment) = False Then
            '    'show failed
            '    MsgBox "Enrollment entry with Enrollment ID : '" & curEnrollment.CN & "' does not exist.", vbExclamation
            '    'close this form
            '    Unload Me
            '    Exit Sub
            'End If
            
            'set form ui info
            With curEnrollment
      
            End With

            'disable some controls
            'txtCN.Enabled = False
            

            
    End Select
    
End Sub


Private Sub Form_Load()
    
    isON = False
        
    'set Customer list
    With b8DPStudent
        Set .DropDBCon = PrimeData
        .SQLFields = "String(10-Len(Trim([StudentID])),'0') & [StudentID] AS CStudentID, (tblStudent.FirstName & ' ' & tblStudent.MiddleName & ' ' & tblStudent.LastName) as FullName, tblStudent.YL"
        .SQLTable = "tblStudent"
        .SQLWhereFields = " CStr(tblStudent.StudentID) & ' ' & tblStudent.FirstName & ' ' & tblStudent.MiddleName & ' ' & tblStudent.LastName "
        .SQLWhere = " tblStudent.Active = True "
        .SQLOrderBy = "tblStudent.LastName, tblStudent.FirstName"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 1
        .AddColumn "Student ID", 100
        .AddColumn "Name", 300
        .AddColumn "Y.L.", 50
    End With
    
End Sub


Private Sub Refresh_YL(Optional bYL As Byte = 0)
    
    With cmbYL
        
        .AddItem "I"
        .AddItem "II"
        .AddItem "III"
        .AddItem "IV"
        .AddItem "V"
        .AddItem "VI"
        
        If bYL = 0 Then
            'select first item
            .ListIndex = 0
        Else
            'select iSem
            .ListIndex = bYL
        End If
        
    End With

End Sub


Private Sub Refresh_SY(Optional iSYID As Integer = 0)
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim ci As ComboItem
    
    
    'disable SY Combo
    cmbSY.Enabled = False
    
    sSQL = "SELECT tblSY.SYID, tblSY.SYTitle" & _
            " From tblSY" & _
            " ORDER BY tblSY.SYID"

    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "Refresh_SY", "ConnectRS(PrimeData, VRS, sSQL) = False / SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    While vRS.EOF = False
        
        cmbSY.ComboItems.Add , "syid//" & ReadField(vRS.Fields("SYID")), _
                            ReadField(vRS.Fields("SYTitle"))
        
        vRS.MoveNext
    Wend
    
    'enable control
    cmbSY.Enabled = True
    
    If iSYID = 0 Then
        'select first item
        cmbSY.ComboItems(1).Selected = True
        GoTo RAE
    End If
    
    For Each ci In cmbSY.ComboItems
        If Val(GetKeyOnSplit(ci.Key, "//", 1)) = iSYID Then
            ci.Selected = True
            Exit For
        End If
    Next
    
RAE:
    Set vRS = Nothing
    Set ci = Nothing
End Sub



Private Sub Refresh_Course(Optional iCourseID As Integer = 0)
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim ci As ComboItem
    
    
    'disable Course Combo
    cmbCourse.Enabled = False
    
    sSQL = "SELECT tblCourse.CourseID, tblCourse.Title" & _
            " From tblCourse" & _
            " ORDER BY tblCourse.Title"


    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "Refresh_Course", "ConnectRS(PrimeData, VRS, sSQL) = False / SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    While vRS.EOF = False
        
        cmbCourse.ComboItems.Add , "Courseid//" & ReadField(vRS.Fields("CourseID")), _
                            ReadField(vRS.Fields("Title"))
        
        vRS.MoveNext
    Wend
    
    'enable control
    cmbCourse.Enabled = True
    
    If iCourseID = 0 Then
        'select first item
        cmbCourse.ComboItems(1).Selected = True
        GoTo RAE
    End If
    
    For Each ci In cmbCourse.ComboItems
        If Val(GetKeyOnSplit(ci.Key, "//", 1)) = iCourseID Then
            ci.Selected = True
            Exit For
        End If
    Next
    
RAE:
    Set vRS = Nothing
    Set ci = Nothing
End Sub

Private Sub Refresh_Sem(Optional iSem As Byte = 0)

    With cmbSem
        .ComboItems.Clear
        .ComboItems.Add , , "1"
        .ComboItems.Add , , "2"
        .ComboItems.Add , , "3"
        
        If iSem = 0 Then
            'select first item
            .ComboItems(1).Selected = True
        Else
            'select iSem
            .ComboItems(iSem).Selected = True
        End If
    End With
End Sub



