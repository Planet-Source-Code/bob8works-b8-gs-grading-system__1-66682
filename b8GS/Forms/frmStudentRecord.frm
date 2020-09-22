VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmStudentRecord 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Grades"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStudentRecord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   ShowInTaskbar   =   0   'False
   Begin b8Controls4.b8TitleBar b8TB 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   661
      Caption         =   "Manage Student's Records/Grades"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColor     =   0
      BorderColor     =   4210752
      BackColor       =   8421504
   End
   Begin VB.PictureBox bgFooter 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   414
      TabIndex        =   0
      Top             =   4440
      Width           =   6210
      Begin VB.Label lblRecSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   585
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   345
         Index           =   0
         Left            =   30
         Top             =   0
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   609
         Color1          =   16777215
         Color2          =   16777215
         Color3          =   14737632
         Color4          =   14737632
         BackColor       =   16119285
      End
   End
   Begin VB.PictureBox bgHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   0
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   530
      TabIndex        =   2
      Top             =   420
      Width           =   7950
      Begin VB.CommandButton cmdDelEnroll 
         Caption         =   "&Delete Enrollment Entry"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1050
         TabIndex        =   13
         Top             =   780
         Width           =   1905
      End
      Begin VB.CommandButton cmdEnroll 
         Caption         =   "&Enroll"
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   780
         Width           =   945
      End
      Begin VB.CheckBox chkAutoRefresh 
         BackColor       =   &H00F5F5F5&
         Caption         =   "&Auto Refresh when changing criteria"
         Height          =   225
         Left            =   2880
         TabIndex        =   9
         Top             =   450
         Value           =   1  'Checked
         Width           =   2955
      End
      Begin VB.CheckBox chkIncludeNonActive 
         BackColor       =   &H00F5F5F5&
         Caption         =   "&Include Non-Active"
         Height          =   225
         Left            =   840
         TabIndex        =   6
         Top             =   450
         Width           =   1695
      End
      Begin b8Controls4.b8DataPicker b8DPStudent 
         Height          =   345
         Left            =   840
         TabIndex        =   7
         Top             =   90
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   609
         SQLWhereSeparator=   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   375
         Index           =   2
         Left            =   90
         Top             =   750
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         Color1          =   16777215
         Color2          =   16777215
         Color3          =   14737632
         Color4          =   14737632
         BackColor       =   16119285
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Student:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   120
         Width           =   630
      End
      Begin b8Controls4.b83DRect shpFooter 
         Height          =   705
         Index           =   1
         Left            =   210
         Top             =   30
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1244
         Color1          =   16777215
         Color2          =   16777215
         Color3          =   14737632
         Color4          =   14737632
         BackColor       =   16119285
      End
   End
   Begin VB.PictureBox bgCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3765
      Left            =   750
      ScaleHeight     =   251
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   1
      Top             =   1380
      Width           =   5760
      Begin VB.PictureBox bgMid 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F5F5F5&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   2445
         Left            =   1530
         MousePointer    =   9  'Size W E
         ScaleHeight     =   163
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   7
         TabIndex        =   11
         Top             =   330
         Width           =   105
      End
      Begin MSComctlLib.ImageList ilList 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStudentRecord.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin b8Controls4.LynxGrid3 listEntries 
         Height          =   2355
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4154
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorBkg    =   16056319
         BackColorSel    =   8438015
         ForeColorSel    =   0
         GridColor       =   11136767
         FocusRectColor  =   33023
         AllowUserResizing=   4
         Striped         =   -1  'True
         SBackColor1     =   16056319
         SBackColor2     =   14940667
      End
      Begin b8Controls4.LynxGrid3 listEnrollment 
         Height          =   2355
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   4154
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorBkg    =   16056319
         BackColorSel    =   8438015
         ForeColorSel    =   0
         GridColor       =   11136767
         FocusRectColor  =   33023
         AllowUserResizing=   4
         Striped         =   -1  'True
         SBackColor1     =   16056319
         SBackColor2     =   14940667
      End
      Begin b8Controls4.b83DRect shpLBorder 
         Height          =   3015
         Left            =   0
         Top             =   30
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   5318
         Color1          =   16777215
         Color2          =   16777215
         Color3          =   14737632
         Color4          =   14737632
         BackColor       =   16119285
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Modify"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuS01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuEnrollment 
      Caption         =   "&Enrollment"
      Visible         =   0   'False
      Begin VB.Menu mnuNewSubject 
         Caption         =   "&New Subject"
      End
      Begin VB.Menu mnuS02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnroll 
         Caption         =   "&Enroll"
      End
      Begin VB.Menu mnuDeleteEnrollment 
         Caption         =   "&Delete Enrollment Entry"
      End
   End
End
Attribute VB_Name = "frmStudentRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Const bgMidDefColor = &HF5F5F5

Dim CurStudent As tStudent

Dim bReadyToDisplay As Boolean
Dim bFormStarted As Boolean

Dim mMidDown As Boolean
Dim iMidX As Single

Public Function ShowForm(Optional ByVal lStudentID As Long = 0)

    If lStudentID > 0 Then
        GetStudentByID lStudentID, CurStudent
        b8DPStudent.BoundData = CurStudent.StudentID
        b8DPStudent.DisplayData = CurStudent.FirstName & " " & CurStudent.MiddleName & " " & CurStudent.LastName
        GoTo Display
    End If
    
    If bFormStarted = True Then
        modFuncChild.ActivateMDIChildForm Me.Name
        Exit Function
    End If
    bFormStarted = True
    

Display:
    'add form
    mdiMain.AddChild Me
    
    'set display flag
    bReadyToDisplay = True

    'load entries
    Form_Refresh
    
End Function

Private Sub b8DPStudent_BeforeDropDown()
     If chkIncludeNonActive.Value = vbChecked Then
        b8DPStudent.SQLWhere = ""
    Else
        b8DPStudent.SQLWhere = " tblStudent.Active = True "
    End If
End Sub

Private Sub b8DPStudent_Change()
    
    If chkAutoRefresh.Value = vbChecked Then
        Me.Form_Refresh
    End If

End Sub

Private Sub bgMid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bgMid.BackColor = &HC0C0C0
    DoEvents
    mMidDown = True
    iMidX = x
End Sub

Private Sub bgMid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim iNewLeft As Single
    
    If mMidDown = True Then
        iNewLeft = bgMid.Left + (x - iMidX)
         bgMid.Left = iNewLeft
    End If
     
    
End Sub

Private Sub bgMid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If mMidDown = True Then
        If bgMid.Left < 80 Then
            bgMid.Left = 80
        End If
        
        On Error Resume Next
        listEnrollment.Move shpLBorder.Left + 3, shpLBorder.Top + 3, bgMid.Left - 5, shpLBorder.Height - 6
        listEntries.Move bgMid.Left + bgMid.Width, shpLBorder.Top + 3, shpLBorder.Width - (bgMid.Left + bgMid.Width) - 6, shpLBorder.Height - 6
        Err.Clear
        listEnrollment.Refresh
        listEntries.Refresh
    End If
    
    bgMid.BackColor = bgMidDefColor
    mMidDown = False
End Sub

Private Sub cmdDelEnroll_Click()
    Dim sEnrollmentID As String
    
    If listEnrollment.RowCount < 1 Then
        Exit Sub
    End If


    If MsgBox("Are you sure you want to delete this Enrollment?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
        Exit Sub
    End If
    
    'get ID
    sEnrollmentID = listEnrollment.CellText(listEnrollment.Row, 1)

    If DeleteEnrollment(sEnrollmentID) = True Then
        Form_Refresh
    Else
        WriteErrorLog Me.Name, "cmdDelEnroll_Click", "Failed on: 'DeleteStudent(lStudentID) = True'"
    End If

End Sub

Private Sub cmdEnroll_Click()
    
    If GetTxtVal(b8DPStudent.BoundData) > 0 Then
        frmEnrollmentEntry.ShowAdd CLng(b8DPStudent.BoundData)
        Form_Refresh
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'pass keyinfo to mdiMain
    mdiMain.AFForm_KeyDown KeyCode, Shift
End Sub

Private Sub listEnrollment_ItemCountChanged()
    cmdDelEnroll.Enabled = IIf(listEnrollment.RowCount > 0, True, False)
End Sub

Private Sub listEnrollment_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu Me.mnuEnrollment
    End If
End Sub

'----------------------------------------------------------
' Controls Procedures
'----------------------------------------------------------
Private Sub listEntries_DblClick()
    Form_Edit
End Sub

Private Sub listEntries_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        
        Me.PopupMenu Me.mnuAction
    End If
End Sub

Private Sub listEntries_RowColChanged()
    RefreshRecSum
End Sub

'----------------------------------------------------------
' >>> END Controls Procedures
'----------------------------------------------------------







'----------------------------------------------------------
' Form Procedures
'----------------------------------------------------------

Private Sub Form_Activate()
    mdiMain.ActivateChild Me
    
    'expand student detail tab
    mdiMain.b8ST(4).Expanded = True
    
    'refresh picture
    mdiMain.Form_ShowStudentDetail
End Sub

Private Sub Form_Load()
        
    'set list columns
    With listEntries
    
        .Redraw = False
        
        .AddColumn "SYTitle", 100   '0
        .AddColumn "SYID", 0   '1
        .AddColumn "EnrolmentID", 0 '2
        .AddColumn "SubjectID", 0   '3
        .AddColumn "Course", 100   '4
        .AddColumn "Sem", 60   '5
        .AddColumn "Term", 60   '6
        .AddColumn "CN", 90   '7
        .AddColumn "SubjectAbr", 100   '8
        .AddColumn "SubjectTitle", 120   '9
        .AddColumn "Grade", 90, lgAlignCenterCenter '10
        .AddColumn "PreReqCN", 90   '11
        .AddColumn "SchedDay", 80   '12
        .AddColumn "Room", 80   '13
        .AddColumn "TimeIn", 90   '14
        .AddColumn "TimeOut", 90   '15
        .AddColumn "Instructor", 120   '16
        
        '.RowHeightMin = 21
        '.ImageList = ilList
        .Redraw = True
        .Refresh
    End With
    
    
    With listEnrollment
    
        .Redraw = False
                
        .AddColumn "S.Y.", 80   '0
        .AddColumn "EnrollmentID", 0
        .AddColumn "SYID", 0
        .AddColumn "CourseID", 0
        .AddColumn "Course", 64, lgAlignCenterCenter
        .AddColumn "Sem", 30, lgAlignCenterCenter
        
        .AddColumn "Created", 0
        .AddColumn "Modified", 0
        .AddColumn "Remarks", 0
        
        .AddColumn "", 6 'spacer
        .AddColumn "Ave. Grade", 70, lgAlignCenterCenter
        
        .RowHeightMin = 21
        .ImageList = ilList
        .Redraw = True
        .Refresh
    End With
    
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

Private Sub Form_Resize()
    
    Dim i As Integer


    On Error Resume Next
    
    'resize title bar
    b8TB.Move 2, 1, Me.ScaleWidth - 4
    'resize header
    bgHeader.Move 0, b8TB.Top + b8TB.Height, Me.ScaleWidth
    'resize footer
    bgFooter.Move 0, Me.ScaleHeight - bgFooter.Height, Me.ScaleWidth
    
    For i = 0 To shpFooter.UBound
        shpFooter(i).Move 2, shpFooter(i).Top, bgFooter.Width - 4
    Next
    
    
    'resize center
    bgCenter.Move 0, bgHeader.Top + bgHeader.Height, Me.ScaleWidth, bgFooter.Top - (bgHeader.Top + bgHeader.Height)

    'resize list
    shpLBorder.Move 2, 0, bgCenter.Width - 4, bgCenter.Height - 0

    listEnrollment.Move shpLBorder.Left + 3, shpLBorder.Top + 3, listEnrollment.Width, shpLBorder.Height - 6
    bgMid.Move listEnrollment.Left + listEnrollment.Width, shpLBorder.Top + 3, 3, shpLBorder.Height - 6
    listEntries.Move listEnrollment.Left + listEnrollment.Width + 3, shpLBorder.Top + 3, shpLBorder.Width - (listEnrollment.Left + listEnrollment.Width + 3) - 6, shpLBorder.Height - 6
    
    
    Err.Clear
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'remove form
    mdiMain.RemoveChild Me.Name
    'clear flag
    bFormStarted = False
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF3 Or (KeyCode = 107 And Shift = 2) Then
        'F3 | Ctrl + '+' - add
        If Form_CanAdd Then
            Form_Add
        End If
    
    ElseIf KeyCode = vbKeyF2 Or (KeyCode = 13 And Shift = 2) Then
        'F2 | Ctrl + Enter - edit
        If Form_CanEdit Then
            Form_Edit
        End If
        
    ElseIf KeyCode = vbKeyDelete Or (KeyCode = 109 And Shift = 2) Then
        'Del : Ctrl + '-' - delete
        If Form_CanDelete Then
            Form_Delete
        End If
    
    ElseIf KeyCode = vbKeyF5 Then
        'F5 - refresh
        If Form_CanRefresh Then
            Form_Refresh
        End If
        
    ElseIf KeyCode = 83 And Shift = 2 Then
        'Ctrl + S - Search
        If Form_CanSearch Then
            mdiMain.Form_ShowSearch
        End If
    
    ElseIf KeyCode = 68 And Shift = 2 Then
        'Ctrl + D - Date Filter
            mdiMain.Form_ShowDateFilter
    End If
End Sub

'----------------------------------------------------------
' >>> END Form Procedures
'----------------------------------------------------------






'----------------------------------------------------------
' Record Procedures
'----------------------------------------------------------
Private Sub LoadEnrollment()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim il As Long
    
    
    
    'set app mouse icon
    mdiMain.Form_StartBussy
    
    'disable
    cmdEnroll.Enabled = False
    'clear list
    listEnrollment.Redraw = False
    listEnrollment.Clear
    
    If bReadyToDisplay = False Then
        GoTo RAE
    End If

        
    If GetTxtVal(b8DPStudent.BoundData) < 1 Then
        GoTo RAE
    End If
    
    sSQL = "SELECT tblEnrollment.EnrollmentID, tblEnrollment.FK_SYID, tblSY.SYTitle, tblCourse.CourseID, tblCourse.Title, tblEnrollment.Sem, tblEnrollment.CreationDate, tblEnrollment.ModifiedDate, tblEnrollment.Remarks, Avg(tblStudentSubject.Grade) AS AvgOfGrade" & _
            " FROM (tblSY INNER JOIN (tblCourse INNER JOIN tblEnrollment ON tblCourse.CourseID = tblEnrollment.FK_CourseID) ON tblSY.SYID = tblEnrollment.FK_SYID) LEFT JOIN tblStudentSubject ON tblEnrollment.EnrollmentID = tblStudentSubject.FK_EnrollmentID" & _
             " WHERE  tblEnrollment.FK_StudentID=" & b8DPStudent.BoundData & _
            " GROUP BY tblEnrollment.EnrollmentID, tblEnrollment.FK_SYID, tblSY.SYTitle, tblCourse.CourseID, tblCourse.Title, tblEnrollment.Sem, tblEnrollment.CreationDate, tblEnrollment.ModifiedDate, tblEnrollment.Remarks" & _
            " ORDER BY tblEnrollment.FK_SYID, tblEnrollment.Sem"


    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "loadenrollment", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    'there is a selected student
    'enable
    cmdEnroll.Enabled = True
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    vRS.MoveFirst
    While vRS.EOF = False
    
        With listEnrollment
        
            il = .AddItem(ReadField(vRS.Fields("SYTitle")))
            .ItemImage(il) = 1
            .CellText(il, 1) = ReadField(vRS.Fields("EnrollmentID"))
            .CellText(il, 2) = ReadField(vRS.Fields("FK_SYID"))
            .CellText(il, 3) = ReadField(vRS.Fields("CourseID"))
            .CellText(il, 4) = ReadField(vRS.Fields("Title"))
            .CellText(il, 5) = ReadField(vRS.Fields("Sem"))
            .CellText(il, 6) = ReadField(vRS.Fields("CreationDate"))
            .CellText(il, 7) = ReadField(vRS.Fields("ModifiedDate"))
            .CellText(il, 8) = ReadField(vRS.Fields("Remarks"))
            .CellText(il, 10) = FormatNumber(ReadField(vRS.Fields("AvgOfGrade")), 2)
        
        End With

        vRS.MoveNext
    Wend
    
    
    
RAE:
    Set vRS = Nothing
    listEnrollment.Redraw = True
    listEnrollment.Refresh
    'restore mouse pointer
    mdiMain.Form_EndBussy
    
End Sub
Private Sub LoadEntries()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim il As Long
    
    
    
    'set app mouse icon
    mdiMain.Form_StartBussy
    
    'clear list
    listEntries.Redraw = False
    listEntries.Clear
    
    
    
    If bReadyToDisplay = False Then
        GoTo RAE
    End If
    
    
    'set SQL Expression
    Dim sSQL_All As String
    
    Dim sSQL_Student As String
    Dim sSQL_SY As String
    Dim sSQL_Sem As String
    Dim sSQL_Term As String

        
    If GetTxtVal(b8DPStudent.BoundData) < 1 Then
        GoTo RAE
    Else
        sSQL_Student = " tblEnrollment.FK_StudentID=" & b8DPStudent.BoundData & ""
    End If

    If GetTxtVal(mdiMain.b8DPSY.BoundData) > 0 Then
        sSQL_SY = " tblSY.SYID=" & mdiMain.b8DPSY.BoundData & " "
    Else
        sSQL_SY = ""
    End If
    
    If GetTxtVal(mdiMain.cmbSem.ListIndex > 0) > 0 Then
        sSQL_Sem = " tblSubject.Sem=" & mdiMain.cmbSem.Text & " "
    Else
        sSQL_Sem = ""
    End If
    
   
    'combine
    sSQL_All = " WHERE " & sSQL_Student
    sSQL_All = sSQL_All & IIf(Len(sSQL_SY) > 0, " AND " & sSQL_SY, "")
    sSQL_All = sSQL_All & IIf(Len(sSQL_Sem) > 0, " AND " & sSQL_Sem, "")

    
    sSQL = "SELECT tblSY.SYID, tblSY.SYTitle, tblCourse.CourseID, tblCourse.Title, tblSubject.Sem, tblSubject.Term, tblSubject.CN, tblSubject.SubjectAbr, tblSubject.SubjectTitle, tblStudentSubject.Grade, tblSubject.PreReqCN, tblSubject.SchedDay, tblSubject.Room, tblSubject.TimeIn, tblSubject.TimeOut, tblSubject.Instructor, tblStudentSubject.FK_EnrollmentID, tblStudentSubject.FK_SubjectID" & _
            " FROM tblSY INNER JOIN (tblSubject INNER JOIN ((tblCourse INNER JOIN tblEnrollment ON tblCourse.CourseID = tblEnrollment.FK_CourseID) INNER JOIN tblStudentSubject ON tblEnrollment.EnrollmentID = tblStudentSubject.FK_EnrollmentID) ON tblSubject.SubjectID = tblStudentSubject.FK_SubjectID) ON tblSY.SYID = tblEnrollment.FK_SYID" & _
            sSQL_All & " ORDER BY tblSY.SYID, tblSubject.Sem, tblSubject.Term, tblSubject.SubjectTitle"


    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    'add entries to list
    Dim iSem As Byte
    Dim ioldSem As Byte
    Dim iSY As Integer
    Dim ioldSY As Integer
    Dim lBgColor As Long
    
    vRS.MoveFirst
    While vRS.EOF = False
    
        With listEntries
        
            iSY = ReadField(vRS.Fields("SYID"))
            iSem = ReadField(vRS.Fields("Sem"))
            
            If iSem <> ioldSem Or iSY <> ioldSY Then
                il = .AddItem(ReadField(vRS.Fields("SYTitle")))
                .CellText(il, 1) = iSY ' 0   '1
                .CellText(il, 2) = ReadField(vRS.Fields("FK_EnrollmentID")) ' 0   '2
                .CellText(il, 3) = ReadField(vRS.Fields("FK_SubjectID")) ' 0   '3
                
                .CellText(il, 4) = ReadField(vRS.Fields("Title")) ' 100   '4
                .CellText(il, 5) = iSem
                .ItemBackColor(il) = &HA9EEFF
            
            End If
        
            il = .AddItem(ReadField(vRS.Fields("SYTitle")))
            .CellText(il, 1) = iSY
            
            .CellText(il, 2) = ReadField(vRS.Fields("FK_EnrollmentID")) ' 0   '2
            .CellText(il, 3) = ReadField(vRS.Fields("FK_SubjectID")) ' 0   '3
            
            .CellText(il, 4) = ReadField(vRS.Fields("Title")) ' 100   '4
            .CellText(il, 5) = iSem
                
 
            .CellForeColor(il, 0) = .CellBackColor(il, 0)
            .CellForeColor(il, 1) = .CellBackColor(il, 1)
            .CellForeColor(il, 2) = .CellBackColor(il, 2)
            .CellForeColor(il, 3) = .CellBackColor(il, 3)
            .CellForeColor(il, 4) = .CellBackColor(il, 4)
            .CellForeColor(il, 5) = .CellBackColor(il, 5)
                
            .CellText(il, 6) = ReadField(vRS.Fields("Term")) ' 60   '6
            .CellText(il, 7) = ReadField(vRS.Fields("CN")) ' 90   '7
            .CellText(il, 8) = ReadField(vRS.Fields("SubjectAbr")) ' 100   '8
            .CellText(il, 9) = ReadField(vRS.Fields("SubjectTitle")) ' 120   '9
            
            .CellFontBold(il, 10) = True
            .CellText(il, 10) = FormatNumber(ReadField(vRS.Fields("Grade")), 2)
            
            .CellText(il, 11) = ReadField(vRS.Fields("PreReqCN")) ' 90   '11
            .CellText(il, 12) = ReadField(vRS.Fields("SchedDay")) ' 80   '12
            .CellText(il, 13) = ReadField(vRS.Fields("Room")) ' 80   '13
            .CellText(il, 14) = ReadField(vRS.Fields("TimeIn")) ' 90   '14
            .CellText(il, 15) = ReadField(vRS.Fields("TimeOut")) ' 90   '15
            .CellText(il, 16) = ReadField(vRS.Fields("Instructor")) ' 120   '16
                
        End With
        
        ioldSem = iSem
        ioldSY = iSY
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    listEntries.Redraw = True
    listEntries.Refresh
    mdiMain.Form_EndBussy
    
End Sub

Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record " & listEntries.Row + 1 & " of " & listEntries.RowCount
End Sub
'----------------------------------------------------------
' >>> END Record Procedures
'----------------------------------------------------------









'----------------------------------------------------------
' Parent Form Calling Functions
'----------------------------------------------------------

Public Function Form_CanAdd() As Boolean
    
    If listEnrollment.RowCount > 0 Then
        Form_CanAdd = True
    End If

End Function
Public Function Form_Add()
    
    Dim li As Long
    
    
    If listEnrollment.RowCount < 1 Then
        Exit Function
    End If
    
    li = listEnrollment.Row

    frmStudentSubjectEntry.ShowAdd listEnrollment.CellText(li, 1), b8DPStudent.DisplayData, listEnrollment.CellText(li, 2), _
                                        CByte(listEnrollment.CellText(li, 5))

    Form_Refresh
    
End Function

Public Function Form_CanEdit() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanEdit = True
    Else
        Form_CanEdit = False
    End If
End Function

Public Function Form_Edit()

    Dim li As Long
    
    
    If listEntries.RowCount < 1 Then
        Exit Function
    End If
    
    li = listEntries.Row

    frmStudentSubjectEntry.ShowEdit listEntries.CellText(li, 2), b8DPStudent.DisplayData, listEntries.CellText(li, 1), _
                                        CByte(listEntries.CellText(li, 5)), , listEntries.CellText(li, 3)
    
    Form_Refresh

End Function

Public Function Form_CanDelete() As Boolean

End Function

Public Function Form_Delete()
    
End Function

Public Function Form_CanRefresh() As Boolean
    Form_CanRefresh = True
End Function

Public Function Form_Refresh()
    
    LoadEnrollment
    LoadEntries
    
    'refresh rec sum
    RefreshRecSum
    'refresh recopt buttons
    mdiMain.ActivateChild Me
    'restore mouse pointer
    mdiMain.Form_EndBussy
    
    'refresh picture
    mdiMain.Form_ShowStudentDetail
End Function

Public Function Form_CanPrint() As Boolean
    Form_CanPrint = False
End Function

Public Function Form_Print()

End Function


Public Function Form_CanSearch() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanSearch = True
    End If
End Function

Public Function Form_SetSearch(ByRef sFields() As String)
    If listEntries.RowCount > 0 Then
    
        ReDim sFields(13)

        sFields(0) = "Enrollemnt ID" ' 0 '0
        sFields(1) = "S.Y. ID" ' 0 '1
        sFields(2) = "S.Y." ' 0 '2
        sFields(3) = "Sem" ' 0 '3
        sFields(4) = "Course ID" ' 0 '4
        sFields(5) = "Course" ' 0 '5
        sFields(6) = "Student ID" ' 0 '6
        sFields(7) = "Last Name" ' 0 '7
        sFields(8) = "First Name" ' 0 '8
        sFields(9) = "Middle Name" ' 0 '9
        sFields(10) = "Y.L." ' 0 '10
        sFields(11) = "Created" ' 0 '11
        sFields(12) = "Modified" ' 0 '12
        sFields(13) = "Remarks" ' 0 '13
       
        Form_SetSearch = True

    End If
End Function
Public Function Form_Search(ByVal sSearchWhat As String, ByVal sField As String) As Boolean
    
    Dim li As Long
    Dim lx As Long
    Dim NewSelIndex As Long
    
    
    'default
    NewSelIndex = -1
    Form_Search = False
    
    listEntries.Redraw = False
    
    If listEntries.RowCount < 1 Then
        GoTo RAE
    End If
    
    If LCase(sField) = "all fields" Then
        
        'all fields
        For lx = 0 To listEntries.Cols - 1
            NewSelIndex = listEntries.FindItem(sSearchWhat, lx, lgWith, False)
            If NewSelIndex >= 0 Then
                listEntries.ItemSelected(NewSelIndex) = True
                listEntries.EnsureVisible NewSelIndex
                Exit For
            End If
        Next
        
    Else
    
        'by column
        For lx = 0 To listEntries.Cols - 1
            If LCase(sField) = LCase(listEntries.ColHeading(lx)) Then
                NewSelIndex = listEntries.FindItem(sSearchWhat, lx, lgWith, False)
                If NewSelIndex >= 0 Then
                    listEntries.ItemSelected(NewSelIndex) = True
                    listEntries.EnsureVisible NewSelIndex
                    
                    Exit For
                End If
                
                Exit For
            End If
        Next
    End If
    
RAE:
    If listEntries.SelectedCount > 1 Then
        For li = 0 To listEntries.RowCount - 1
            If NewSelIndex <> li Then
            listEntries.ItemSelected(li) = False
            End If
        Next
    End If
    
    listEntries.Redraw = True
    listEntries.Refresh
    
    'return
    If NewSelIndex >= 0 Then
        Form_Search = True
    End If
    
End Function


Public Function Form_CanFilter_SY() As Boolean
    Form_CanFilter_SY = True
End Function
Public Function Form_CanFilter_Course() As Boolean
    Form_CanFilter_Course = True
End Function
Public Function Form_CanFilter_Sem() As Boolean
    Form_CanFilter_Sem = True
End Function
Public Function Form_CanFilter_Term() As Boolean
    Form_CanFilter_Term = False
End Function

Public Sub Form_FilterChange()
    
    If chkAutoRefresh.Value = vbChecked Then
        Me.Form_Refresh
    End If
    
End Sub


Public Function Form_CanShowStudentDetail() As Boolean
    'If GetTxtVal(b8DPStudent.BoundData) > 0 Then
        Form_CanShowStudentDetail = True
    'End If
End Function

Public Function Form_GetStudentID() As Long
    
    On Error GoTo RAE
    
    Form_GetStudentID = GetTxtVal(b8DPStudent.BoundData)
    
    Exit Function
RAE:
End Function





'----------------------------------------------------------
' Menu Procedures
'----------------------------------------------------------


Private Sub mnuAction_Click()
    mnuAdd.Enabled = Form_CanAdd
    mnuEdit.Enabled = Form_CanEdit
    mnuDelete.Enabled = Form_CanDelete
    mnuRefresh.Enabled = Form_CanRefresh
    mnuPrint.Enabled = Form_CanPrint
    mnuSearch.Enabled = Form_CanSearch
End Sub

Private Sub mnuAdd_Click()
    Form_Add
End Sub


Private Sub mnuDelete_Click()
    Form_Delete
End Sub

Private Sub mnuDeleteEnrollment_Click()
    cmdDelEnroll_Click
End Sub

Private Sub mnuEdit_Click()
    Form_Edit
End Sub

Private Sub mnuEnroll_Click()
    cmdEnroll_Click
End Sub

Private Sub mnuEnrollment_Click()

    If GetTxtVal(b8DPStudent.BoundData) > 0 Then
        mnuEnroll.Enabled = True
    Else
        mnuEnroll.Enabled = False
    End If
    
    If listEnrollment.RowCount > 0 Then
        mnuNewSubject.Enabled = True
        mnuDeleteEnrollment.Enabled = True
    Else
        mnuNewSubject.Enabled = False
        mnuDeleteEnrollment.Enabled = False
    End If

End Sub

Private Sub mnuNewSubject_Click()
    Form_Add
End Sub

Private Sub mnuPrint_Click()
    Form_Print
End Sub

Private Sub mnuRefresh_Click()
    Form_Refresh
End Sub

Private Sub mnuSearch_Click()
    
    mdiMain.Form_ShowSearch
End Sub

'----------------------------------------------------------
' >>> END Menu Procedures
'----------------------------------------------------------
