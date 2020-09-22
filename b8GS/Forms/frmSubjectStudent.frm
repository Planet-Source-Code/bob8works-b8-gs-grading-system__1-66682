VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmSubjectStudent 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Ranking"
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
   Icon            =   "frmSubjectStudent.frx":0000
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
      Caption         =   "Manage Ranking List"
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
      Begin VB.TextBox txtMaxCount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Text            =   "100"
         Top             =   780
         Width           =   795
      End
      Begin VB.CheckBox chkAutoRefresh 
         BackColor       =   &H00F5F5F5&
         Caption         =   "&Auto Refresh when changing criteria"
         Height          =   225
         Left            =   840
         TabIndex        =   8
         Top             =   450
         Value           =   1  'Checked
         Width           =   2955
      End
      Begin b8Controls4.b8DataPicker b8DPSubject 
         Height          =   345
         Left            =   840
         TabIndex        =   6
         Top             =   90
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
         DropWinWidth    =   8735
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "record/s only."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2430
         TabIndex        =   10
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Show Top"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   810
         TabIndex        =   9
         Top             =   810
         Width           =   705
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
         Caption         =   "&Subject:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   120
         Width           =   600
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
      Left            =   720
      ScaleHeight     =   251
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   1
      Top             =   1380
      Width           =   5760
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
               Picture         =   "frmSubjectStudent.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin b8Controls4.LynxGrid3 listEntries 
         Height          =   2355
         Left            =   330
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
Attribute VB_Name = "frmSubjectStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim bReadyToDisplay As Boolean
Dim bFormStarted As Boolean

Dim mMidDown As Boolean
Dim iMidX As Single

Dim RSRANK As New ADODB.Recordset

Public Function ShowForm()


    If bFormStarted = True Then
        modFuncChild.ActivateMDIChildForm Me.Name
        Exit Function
    End If
    bFormStarted = True

    'add form
    mdiMain.AddChild Me
    
    'set display flag
    bReadyToDisplay = True

    'load entries
    LoadEntries
    
End Function



Private Sub b8DPSubject_Change()

    If chkAutoRefresh.Value = vbChecked Then
        Me.Form_Refresh
    End If

    If GetTxtVal(b8DPSubject.BoundData) < 1 Then
        b8DPSubject.DisplayData = "*by criteria"
    End If
End Sub







Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'pass keyinfo to mdiMain
    mdiMain.AFForm_KeyDown KeyCode, Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set RSRANK = Nothing
End Sub

'----------------------------------------------------------
' Controls Procedures
'----------------------------------------------------------
Private Sub listEntries_DblClick()
    If listEntries.RowCount > 0 Then
        frmStudentRecord.ShowForm CLng(listEntries.CellText(listEntries.Row, 1))
    End If
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
        
        .AddColumn "Enrollment ID", 0   '0
        .AddColumn "Student ID", 100   '1
        .AddColumn "Name", 240   '2
        .AddColumn "Course", 70   '3
        .AddColumn "Y.L.", 40   '4
        .AddColumn "Grade", 80, lgAlignRightCenter   '5
        
        
        '.RowHeightMin = 21
        '.ImageList = ilList
        .Redraw = True
        .Refresh
    End With
    
     'set Customer list
    With b8DPSubject
        Set .DropDBCon = PrimeData
        .SQLFields = "tblSubject.SubjectID, tblSubject.FK_SYID, tblSY.SYTitle, tblSubject.Sem, tblSubject.Term, tblSubject.CN, tblSubject.SubjectAbr, tblSubject.SubjectTitle, tblSubject.PreReqCN, tblSubject.TimeIn, tblSubject.TimeOut, tblSubject.Room, tblSubject.SchedDay, tblSubject.Instructor"
        .SQLTable = "tblSY INNER JOIN tblSubject ON tblSY.SYID = tblSubject.FK_SYID"
        .SQLWhereFields = " CStr(tblSubject.SubjectID) & ' ' & CStr(tblSubject.FK_SYID) & ' ' & tblSY.SYTitle & ' ' & CStr(tblSubject.Sem) & ' ' & CStr(tblSubject.Term) & ' ' & CStr(tblSubject.CN) & ' ' & tblSubject.SubjectAbr & ' ' & tblSubject.SubjectTitle & ' ' & CStr(tblSubject.PreReqCN) & ' ' & CStr(tblSubject.TimeIn) & ' ' & CStr(tblSubject.TimeOut) & ' ' & tblSubject.Room & ' ' & tblSubject.SchedDay & ' ' & tblSubject.Instructor "
        .SQLOrderBy = "tblSubject.FK_SYID, tblSubject.Sem, tblSubject.Term, tblSubject.SubjectAbr"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 7
        .AddColumn "SubjectID", 100    '0
        .AddColumn "FKSYID", 0    '1
        .AddColumn "SYTitle", 80    '2
        .AddColumn "Sem", 50    '3
        .AddColumn "Term", 50    '4
        .AddColumn "CN", 70    '5
        .AddColumn "SubjectAbr", 70    '6
        .AddColumn "SubjectTitle", 100    '7
        .AddColumn "PreReqCN", 0    '8
        .AddColumn "TimeOut", 70    '9
        .AddColumn "Room", 70    '10
        .AddColumn "SchedDay", 60    '11
        .AddColumn "Instructor", 120    '12

    End With
    
    'SELECT tblSubject.SubjectID, tblSubject.FK_SYID, tblSY.SYTitle, tblSubject.Sem, tblSubject.Term, tblSubject.CN, tblSubject.SubjectAbr, tblSubject.SubjectTitle, tblSubject.PreReqCN, tblSubject.TimeIn, tblSubject.TimeOut, tblSubject.Room, tblSubject.SchedDay, tblSubject.Instructor
    'FROM tblSY INNER JOIN tblSubject ON tblSY.SYID = tblSubject.FK_SYID
    'ORDER BY tblSubject.FK_SYID, tblSubject.Sem, tblSubject.Term, tblSubject.SubjectAbr;

    
    
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

    listEntries.Redraw = False
    listEntries.Move shpLBorder.Left + 3, shpLBorder.Top + 3, shpLBorder.Width - 6, shpLBorder.Height - 6
    listEntries.Redraw = True
    listEntries.Refresh
    
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

Private Sub LoadEntries()
    
    
    Dim sSQL As String
    Dim il As Long
    
    Dim iMaxCount As Integer
    Dim ic As Integer
    
    
    'set app mouse icon
    mdiMain.Form_StartBussy
    
    'clear list
    listEntries.Redraw = False
    listEntries.Clear
    
    
    
    If bReadyToDisplay = False Then
        GoTo RAE
    End If
    
    'ceck max record count to be display
    iMaxCount = GetTxtVal(txtMaxCount.Text)
    If iMaxCount < 1 Or iMaxCount > 1000 Then
        MsgBox "Please enter valid Record count to be display.", vbExclamation
        HLTxt txtMaxCount
        Exit Sub
    End If
    
    txtMaxCount.Text = iMaxCount
    

    If GetTxtVal(b8DPSubject.BoundData) > 0 Then
        sSQL = "SELECT TOP " & iMaxCount & " tblEnrollment.EnrollmentID, (String(10-Len(Trim(tblEnrollment.FK_StudentID)),'0')&tblEnrollment.FK_StudentID) as StudIDFull, (tblStudent.Lastname & ', ' & tblStudent.FirstName & ' ' & tblStudent.MiddleName) AS StudentFullName, tblCourse.Title, Choose(tblEnrollment.YL,'I','II','III','IV','V') as TYL, (tblStudentSubject.Grade) AS AVGOfGrade" & _
            " FROM tblCourse INNER JOIN (tblStudent INNER JOIN (tblEnrollment INNER JOIN tblStudentSubject ON tblEnrollment.EnrollmentID = tblStudentSubject.FK_EnrollmentID) ON tblStudent.StudentID = tblEnrollment.FK_StudentID) ON tblCourse.CourseID = tblEnrollment.FK_CourseID" & _
            " Where (((tblStudentSubject.FK_SubjectID) = '" & b8DPSubject.BoundData & "' ))" & _
            " ORDER BY tblStudentSubject.Grade DESC"
            

     Else
     
        'set SQL Expression
        Dim sSQL_All As String
        
        Dim sSQL_Student As String
        Dim sSQL_SY As String
        Dim sSQL_Course As String
        Dim sSQL_Sem As String
        Dim sSQL_Term As String
        Dim sSQL_YL As String
     
        'Where (((tblEnrollment.FK_SYID) = 1) And ((tblEnrollment.FK_CourseID) = 1) And ((tblEnrollment.Sem) = 1))
    
        If GetTxtVal(mdiMain.b8DPSY.BoundData) > 0 Then
            sSQL_SY = " tblEnrollment.FK_SYID=" & mdiMain.b8DPSY.BoundData & " "
        Else
            sSQL_SY = ""
        End If
        
        
        If GetTxtVal(mdiMain.b8DPCourse.BoundData) > 0 Then
            sSQL_Course = " tblEnrollment.FK_CourseID=" & mdiMain.b8DPCourse.BoundData & " "
        Else
            sSQL_Course = ""
        End If
        
        If mdiMain.cmbSem.ListIndex > 0 Then
            sSQL_Sem = " tblEnrollment.Sem=" & mdiMain.cmbSem.Text & " "
        Else
            sSQL_Sem = ""
        End If
        
        If mdiMain.cmbYL.ListIndex > 0 Then
            sSQL_YL = " tblEnrollment.YL=" & modYL.YLStrToNum(mdiMain.cmbYL) & " "
        Else
            sSQL_YL = ""
        End If
              
        
        'combine

        sSQL_All = " WHERE tblStudent.Active=True" & IIf(Len(sSQL_SY) > 0, " AND " & sSQL_SY, "") '
        sSQL_All = sSQL_All & IIf(Len(sSQL_Course) > 0, " AND " & sSQL_Course, "")
        sSQL_All = sSQL_All & IIf(Len(sSQL_Sem) > 0, " AND " & sSQL_Sem, "")
        sSQL_All = sSQL_All & IIf(Len(sSQL_YL) > 0, " AND " & sSQL_YL, "")

        
        sSQL = "SELECT TOP " & iMaxCount & " tblEnrollment.EnrollmentID, (String(10-Len(Trim(tblEnrollment.FK_StudentID)),'0')&tblEnrollment.FK_StudentID) as StudIDFull, (tblStudent.LastName & ', ' & tblStudent.FirstName & ' ' & tblStudent.MiddleName) AS StudentFullName, tblCourse.Title, Choose(tblEnrollment.YL,'I','II','III','IV','V') as TYL, Avg(tblStudentSubject.Grade) AS AvgOfGrade" & _
                " FROM tblStudent INNER JOIN ((tblCourse INNER JOIN tblEnrollment ON tblCourse.CourseID = tblEnrollment.FK_CourseID) INNER JOIN tblStudentSubject ON tblEnrollment.EnrollmentID = tblStudentSubject.FK_EnrollmentID) ON tblStudent.StudentID = tblEnrollment.FK_StudentID" & _
                sSQL_All & _
                " GROUP BY tblEnrollment.EnrollmentID, (String(10-Len(Trim(tblEnrollment.FK_StudentID)),'0')&tblEnrollment.FK_StudentID), (tblStudent.LastName & ', ' & tblStudent.FirstName & ' ' & tblStudent.MiddleName), tblCourse.Title, tblEnrollment.YL" & _
                " ORDER BY Avg(tblStudentSubject.Grade) DESC"
    End If

    
    If ConnectRS(PrimeData, RSRANK, sSQL) = False Then
        WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(RSRANK) = False Then
        GoTo RAE
    End If
    
    
    RSRANK.MoveFirst
    ic = 0
    While RSRANK.EOF = False
    
        With listEntries
            
            il = .AddItem(ReadField(RSRANK.Fields("EnrollmentID")))
            .CellText(il, 1) = ReadField(RSRANK.Fields("StudIDFull"))
            .CellFontBold(il, 2) = True
            .CellText(il, 2) = ic + 1 & ".    " & ReadField(RSRANK.Fields("StudentFullName"))
            .CellText(il, 3) = ReadField(RSRANK.Fields("Title"))
            .CellText(il, 4) = ReadField(RSRANK.Fields("TYL"))
            .CellFontBold(il, 5) = True
            .CellText(il, 5) = FormatNumber(ReadField(RSRANK.Fields("AvgOfGrade")), 2)
        
        End With

        ic = ic + 1
        If ic >= iMaxCount Then
            GoTo RAE
        End If
        RSRANK.MoveNext
    Wend
    
RAE:
    listEntries.Redraw = True
    listEntries.Refresh
    mdiMain.Form_EndBussy
    
End Sub

Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record " & listEntries.Row + 1 & " of " & listEntries.RowCount

    mdiMain.Form_ShowStudentDetail
    
End Sub
'----------------------------------------------------------
' >>> END Record Procedures
'----------------------------------------------------------









'----------------------------------------------------------
' Parent Form Calling Functions
'----------------------------------------------------------

Public Function Form_CanAdd() As Boolean


End Function
Public Function Form_Add()

    
End Function

Public Function Form_CanEdit() As Boolean

End Function

Public Function Form_Edit()

End Function

Public Function Form_CanDelete() As Boolean

End Function

Public Function Form_Delete()
    
End Function

Public Function Form_CanRefresh() As Boolean
    Form_CanRefresh = True
End Function

Public Function Form_Refresh()

    LoadEntries
    
    'refresh rec sum
    RefreshRecSum
    'refresh recopt buttons
    mdiMain.ActivateChild Me
    'restore mouse pointer
    
End Function

Public Function Form_CanPrint() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanPrint = True
    End If
End Function

Public Function Form_Print()
    drRank.ShowReport RSRANK
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
    If listEntries.RowCount > 0 Then
        Form_CanShowStudentDetail = True
    End If
End Function

Public Function Form_GetStudentID() As Long
    
    On Error GoTo RAE
    If listEntries.RowCount > 0 Then
        Form_GetStudentID = Val(listEntries.CellText(listEntries.Row, 0))
    End If
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
