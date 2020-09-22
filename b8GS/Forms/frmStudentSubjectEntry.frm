VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStudentSubjectEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading..."
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStudentSubjectEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
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
      TabIndex        =   20
      Top             =   0
      Width           =   10305
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student's Subject Entry"
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
         TabIndex        =   21
         Top             =   90
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmStudentSubjectEntry.frx":000C
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdRefreshSubjects 
      Caption         =   "&Refresh"
      Height          =   315
      Left            =   8400
      TabIndex        =   19
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddSubject 
      Caption         =   "&New Subject"
      Height          =   315
      Left            =   7200
      TabIndex        =   18
      Top             =   3060
      Width           =   1185
   End
   Begin VB.TextBox txtRemarks 
      Height          =   855
      Left            =   4320
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2040
      Width           =   5025
   End
   Begin VB.TextBox txtGrade 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   13
      Top             =   2490
      Width           =   1500
   End
   Begin MSComctlLib.ListView listsubjects 
      Height          =   2355
      Left            =   120
      TabIndex        =   12
      Top             =   3390
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CN"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Day"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Room"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Instructor"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Grade"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Remarks"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "SortKey"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.ComboBox cmbSem 
      BackColor       =   &H00E3FFFE&
      Height          =   315
      Left            =   4290
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1560
      Width           =   795
   End
   Begin VB.ComboBox cmbTerm 
      BackColor       =   &H00E3FFFE&
      Height          =   315
      Left            =   6210
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtSY 
      BackColor       =   &H00E3FFFE&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1590
      Width           =   1515
   End
   Begin VB.TextBox txtStudentFullName 
      BackColor       =   &H00E3FFFE&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   810
      Width           =   7755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   8220
      TabIndex        =   3
      Top             =   6060
      Width           =   1395
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Update"
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
      Left            =   6660
      TabIndex        =   2
      Top             =   6060
      Width           =   1395
   End
   Begin VB.TextBox txtCN 
      BackColor       =   &H00E3FFFE&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2040
      Width           =   1500
   End
   Begin b8Controls4.b8Line b8Line1 
      Height          =   30
      Left            =   0
      TabIndex        =   22
      Top             =   555
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   53
      BorderColor1    =   15592425
      BorderColor2    =   16777215
   End
   Begin b8Controls4.b8Line b8Line2 
      Height          =   30
      Left            =   -240
      TabIndex        =   23
      Top             =   5910
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   53
      BorderColor1    =   15592425
      BorderColor2    =   16777215
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Subjects:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   3150
      Width           =   1365
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3420
      TabIndex        =   16
      Top             =   2250
      Width           =   675
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   14
      Top             =   2490
      Width           =   495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Term:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5490
      TabIndex        =   11
      Top             =   1590
      Width           =   420
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sem:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3630
      TabIndex        =   8
      Top             =   1590
      Width           =   360
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S.Y.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   7
      Top             =   1620
      Width           =   360
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   870
      Width           =   630
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CN:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   2070
      Width           =   270
   End
End
Attribute VB_Name = "frmStudentSubjectEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mFormState As String

Dim curStudentSubject As tStudentSubject

Dim curStudentFullName As String
Dim curSYID As Integer
Dim curSem As Byte
Dim curTerm As Byte

Dim curSubjectIDOnEdit As String


Dim mShowAdd As Boolean
Dim mShowEdit As Boolean

Dim isON As Boolean

Public Function ShowAdd(ByVal sEnrollmentID As String, _
                        ByVal sStudentFullName As String, _
                        ByVal iSYID As Integer, _
                        ByVal bSem As Byte, _
                        Optional ByVal bTerm As Byte = 1, _
                        Optional ByVal sSubjectID As String = "") As Boolean
    
    'set form state
    mFormState = "add"
    
    curStudentSubject.FK_EnrollmentID = sEnrollmentID
    curStudentSubject.FK_SubjectID = sSubjectID
    
        
    curStudentFullName = sStudentFullName
    curSYID = iSYID
    curSem = bSem
    curTerm = bTerm

    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowEdit(ByVal sEnrollmentID As String, _
                        ByVal sStudentFullName As String, _
                        ByVal iSYID As Integer, _
                        ByVal bSem As Byte, _
                        Optional ByVal bTerm As Byte = 1, _
                        Optional ByVal sSubjectID As String = "") As Boolean
    
    'set form state
    mFormState = "edit"
    
    
    curStudentSubject.FK_EnrollmentID = sEnrollmentID
    curStudentSubject.FK_SubjectID = sSubjectID
    
    curSubjectIDOnEdit = sSubjectID
    
    curStudentFullName = sStudentFullName
    curSYID = iSYID
    curSem = bSem
    curTerm = bTerm
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function



Private Sub cmbTerm_Click()
    
    'change current Term
    curTerm = GetTxtVal(cmbTerm.Text)
    'refresh available subjects
    RefershSubjects
    'check all enrolled subjects
    RefreshStudentSubject
    
    'set active subject
    listSubjects_DblClick
End Sub

Private Sub cmdAddSubject_Click()
    Dim iSYID As Integer
    Dim bSem As Byte
    Dim bTerm As Byte
    
    On Error Resume Next
    
    iSYID = CurrentSY.SYID
    bSem = 1
    bTerm = 1
    iSYID = CInt(Left(txtSY.Text, 4))
    bSem = CByte(cmbSem.Text)
    bTerm = CByte(cmbTerm.Text)

    frmSubjectEntry.ShowAdd iSYID, bSem, bTerm
    Err.Clear
 
    RefershSubjects
    RefreshStudentSubject
    
   
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefreshSubjects_Click()
    RefreshStudentSubject
End Sub

Private Sub cmdSave_Click()

    'validate
    If GetTxtVal(txtGrade.Text) < 60 Or GetTxtVal(txtGrade.Text) > 100 Then
        MsgBox "Please enter valid grade." & vbNewLine & vbNewLine & "Range: 60-100", vbExclamation
        HLTxt txtGrade
        Exit Sub
    End If

    If listsubjects.ListItems.Count < 1 Then
        Exit Sub
    End If
    
    If listsubjects.SelectedItem.Checked = True Then
        SaveEdit
    Else
        SaveAdd
    End If
    
End Sub

Private Function SaveEdit()

    
    If IsEmpty(txtCN.Text) = True Then
        MsgBox "Please select subject first.", vbExclamation
        Exit Function
    End If
    
    curStudentSubject.ModifiedDate = Now
    curStudentSubject.Grade = Val(txtGrade.Text)
    curStudentSubject.Remarks = txtRemarks.Text
    
    If EditStudentSubject(curStudentSubject) = True Then
        'success
        'check all enrolled subjects
        RefreshStudentSubject
        'set flag
        mShowEdit = True
    Else
        'failed
        WriteErrorLog Me.Name, "SaveEdit", "AddStudentSubject(curStudentSubject) = False"
    End If
    
End Function

Private Function SaveAdd()
   
    If IsEmpty(txtCN.Text) = True Then
        MsgBox "Please select subject first.", vbExclamation
        Exit Function
    End If
    
    curStudentSubject.ModifiedDate = Now
    curStudentSubject.Grade = Val(txtGrade.Text)
    curStudentSubject.Remarks = txtRemarks.Text
    
    If AddStudentSubject(curStudentSubject) = True Then
        'success
        'check all enrolled subjects
        RefreshStudentSubject
        'set flag
        mShowAdd = True
    Else
        'failed
        WriteErrorLog Me.Name, "SaveAdd", "AddStudentSubject(curStudentSubject) = False"
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
            Me.Caption = "New Student's Grade"

            'set form fields
            txtStudentFullName.Text = curStudentFullName
            txtSY.Text = curSYID & "-" & curSYID + 1
            cmbSem.ListIndex = curSem - 1
            cmbTerm.ListIndex = curTerm - 1
            
            
            
        Case "edit"

            'set caption
            Me.Caption = "Edit Student's Grade"
            
            'set form fields
            txtStudentFullName.Text = curStudentFullName
            txtSY.Text = curSYID & "-" & curSYID + 1
            cmbSem.ListIndex = curSem - 1
            cmbTerm.ListIndex = curTerm - 1
            
            
    
    End Select
    
End Sub


Private Sub Form_Load()
    
    'set sem list items
    cmbSem.Clear
    cmbSem.AddItem "1"
    cmbSem.AddItem "2"
    cmbSem.AddItem "3"
    'set term list items
    cmbTerm.Clear
    cmbTerm.AddItem "All"
    cmbTerm.AddItem "1"
    cmbTerm.AddItem "2"
    cmbTerm.AddItem "3"
    
    
    isON = False
End Sub

Private Sub txtStudentSubjectID_Change()
    
   ' txtStudentSubjectTitle.Text = ""
    
    If IsNumeric(txtStudentSubjectID.Text) = False Then
        Exit Sub
    End If
    
    If Val(txtStudentSubjectID.Text) < 1000 Or Val(txtStudentSubjectID.Text) > 9998 Then
        Exit Sub
    End If
    
    txtStudentSubjectTitle.Text = GetStudentSubjectTitle(Val(txtStudentSubjectID.Text))
    
End Sub



Private Sub RefershSubjects()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    
    txtCN.Text = ""
    txtGrade.Text = ""
    curStudentSubject.FK_SubjectID = ""
    
    listsubjects.Sorted = False
    listsubjects.ListItems.Clear
    
  
    If curTerm > 0 Then
        sSQL = "SELECT tblSubject.SubjectID, tblSubject.CN, tblSubject.SubjectAbr, tblSubject.SubjectTitle, tblSubject.FK_SYID, tblSubject.Sem, tblSubject.Term, tblSubject.PreReqCN, tblSubject.TimeIn, tblSubject.TimeOut, tblSubject.Room, tblSubject.SchedDay, tblSubject.Instructor" & _
                " From tblSubject" & _
                " WHERE tblSubject.FK_SYID=" & curSYID & " AND tblSubject.Sem=" & curSem & " AND tblSubject.Term=" & curTerm & _
                " ORDER BY tblSubject.Sem, tblSubject.Term, tblSubject.SubjectAbr"
    Else
        sSQL = "SELECT tblSubject.SubjectID, tblSubject.CN, tblSubject.SubjectAbr, tblSubject.SubjectTitle, tblSubject.FK_SYID, tblSubject.Sem, tblSubject.Term, tblSubject.PreReqCN, tblSubject.TimeIn, tblSubject.TimeOut, tblSubject.Room, tblSubject.SchedDay, tblSubject.Instructor" & _
                " From tblSubject" & _
                " WHERE tblSubject.FK_SYID=" & curSYID & " AND tblSubject.Sem=" & curSem & _
                " ORDER BY tblSubject.Sem, tblSubject.Term, tblSubject.SubjectAbr"
    End If
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "RefershSubjects", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        listsubjects.Enabled = False
        GoTo RAE
    End If
    
    listsubjects.Enabled = True
       
    vRS.MoveFirst
    
    While vRS.EOF = False
        listsubjects.ListItems.Add , ReadField(vRS.Fields("SubjectID")), ReadField(vRS.Fields("CN"))
        With listsubjects.ListItems.Item(listsubjects.ListItems.Count)
            .SubItems(1) = ReadField(vRS.Fields("SubjectAbr"))
            .SubItems(2) = ReadField(vRS.Fields("SubjectTitle"))
            .SubItems(3) = ReadField(vRS.Fields("SchedDay"))
            .SubItems(4) = ReadField(vRS.Fields("TimeIn")) & "-" & ReadField(vRS.Fields("TimeOut"))
            .SubItems(5) = ReadField(vRS.Fields("Room"))
            .SubItems(6) = ReadField(vRS.Fields("Instructor"))
            .SubItems(9) = "9999999999"
        End With
        'cmbsubject
        vRS.MoveNext
    Wend

RAE:
    Set vRS = Nothing
End Sub
    
Private Sub RefreshCurrentSubject()

    If listsubjects.ListItems.Count > 0 Then
        txtCN.Text = listsubjects.SelectedItem.SubItems(1)
        txtGrade.Text = FormatNumber(Val(listsubjects.SelectedItem.SubItems(7)), 2)
        curStudentSubject.FK_SubjectID = listsubjects.SelectedItem.Key
        txtRemarks.Text = listsubjects.SelectedItem.SubItems(8)
    End If
    
End Sub
Private Sub RefreshStudentSubject()
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As ListItem
    Dim ic As Integer

    Dim sSubjectID As String

    sSQL = "SELECT tblStudentSubject.FK_EnrollmentID, tblStudentSubject.FK_SubjectID, tblStudentSubject.Grade, tblStudentSubject.Remarks" & _
            " FROM tblSubject INNER JOIN tblStudentSubject ON tblSubject.SubjectID = tblStudentSubject.FK_SubjectID" & _
            " WHERE (((tblStudentSubject.FK_EnrollmentID)='" & curStudentSubject.FK_EnrollmentID & "'))"
    
    
    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "RefreshStudentSubject", "Unable to connect Recordset. SQL expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
       
    vRS.MoveFirst
    ic = 0
    While vRS.EOF = False
        sSubjectID = ReadField(vRS.Fields("FK_SubjectID"))
        For Each li In listsubjects.ListItems
        
            If li.Key = sSubjectID Then
                'check if on edit

                If mFormState = "edit" Then

                    If li.Key = curSubjectIDOnEdit Then
                        li.Selected = True
                    End If
                End If

                'found
                li.Checked = True
                li.SubItems(7) = FormatNumber(ReadField(vRS.Fields("Grade")), 2)
                li.SubItems(8) = ReadField(vRS.Fields("Remarks"))
                li.SubItems(9) = modFunction.ComNumZ(ic, 10)
                ic = ic + 1
                Exit For
                
                
                
            End If
        Next
        vRS.MoveNext
    Wend
    
    'sort
    listsubjects.Sorted = False
    listsubjects.SortKey = 9
    listsubjects.SortOrder = lvwAscending
    listsubjects.Sorted = True
    listsubjects.Refresh
    

RAE:
    Set li = Nothing
    Set vRS = Nothing
End Sub

Private Sub listsubjects_Click()
    RefreshCurrentSubject
End Sub

Private Sub listSubjects_DblClick()
    RefreshCurrentSubject
    HLTxt txtGrade
End Sub

Private Sub listsubjects_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        'from false
        Item.Checked = False
    Else
        'from true
        If MsgBox("Are you sure you want to remove '" & Item.SubItems(2) & "'?", vbQuestion + vbOKCancel) = vbOK Then
            'delete student subject
            If DeleteStudentSubject(curStudentSubject.FK_EnrollmentID, curStudentSubject.FK_SubjectID) = True Then
                'success
                'refresh available subjects
                RefershSubjects
                'check all enrolled subjects
                RefreshStudentSubject
            Else
                WriteErrorLog Me.Name, "listsubjects_ItemCheck", "DeleteStudentSubject(curStudentSubject.FK_EnrollmentID, curStudentSubject.FK_SubjectID) = False"
            End If
        Else
            Item.Checked = True
        End If
    End If
End Sub

Private Sub listsubjects_KeyUp(KeyCode As Integer, Shift As Integer)
    RefreshCurrentSubject
End Sub


Private Sub txtGrade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        MsgBox ":"
    End If
End Sub
