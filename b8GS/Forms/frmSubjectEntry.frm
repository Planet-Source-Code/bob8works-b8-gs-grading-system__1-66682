VERSION 5.00
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubjectEntry 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading..."
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSubjectEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSubjectEntry.frx":000C
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   533
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
      TabIndex        =   26
      Top             =   0
      Width           =   10305
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Entry"
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
         TabIndex        =   27
         Top             =   90
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmSubjectEntry.frx":034E
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.TextBox txtSubjectID 
      Height          =   315
      Left            =   1125
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   840
      Width           =   2100
   End
   Begin VB.TextBox txtSchedDay 
      Height          =   315
      Left            =   5685
      TabIndex        =   17
      Top             =   3180
      Width           =   2100
   End
   Begin VB.TextBox txtInstructor 
      Height          =   315
      Left            =   1140
      TabIndex        =   14
      Top             =   4020
      Width           =   6660
   End
   Begin VB.TextBox txtRoom 
      Height          =   315
      Left            =   5685
      TabIndex        =   12
      Top             =   3600
      Width           =   2100
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   315
      Left            =   5685
      TabIndex        =   10
      Top             =   2760
      Width           =   2100
   End
   Begin VB.TextBox txtTimeIn 
      Height          =   315
      Left            =   5685
      TabIndex        =   8
      Top             =   2340
      Width           =   2100
   End
   Begin VB.TextBox txtSubjectAbr 
      Height          =   315
      Left            =   5685
      TabIndex        =   6
      Top             =   1200
      Width           =   2100
   End
   Begin VB.TextBox txtSubjectTitle 
      Height          =   315
      Left            =   1125
      TabIndex        =   4
      Top             =   1620
      Width           =   6660
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   4710
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
      Left            =   4920
      TabIndex        =   2
      Top             =   4710
      Width           =   1395
   End
   Begin VB.TextBox txtCN 
      Height          =   315
      Left            =   1125
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1260
      Width           =   2100
   End
   Begin MSComctlLib.ImageCombo cmbSem 
      Height          =   330
      Left            =   1140
      TabIndex        =   19
      Top             =   3180
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageCombo cmbTerm 
      Height          =   330
      Left            =   1140
      TabIndex        =   21
      Top             =   3600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageCombo cmbSY 
      Height          =   330
      Left            =   1140
      TabIndex        =   25
      Top             =   2340
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
      TabIndex        =   28
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
      TabIndex        =   29
      Top             =   4560
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
      Caption         =   "Subject ID:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   24
      Top             =   900
      Width           =   810
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Term:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   22
      Top             =   3660
      Width           =   420
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
      TabIndex        =   20
      Top             =   3240
      Width           =   420
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Day:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   18
      Top             =   3255
      Width           =   345
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
      TabIndex        =   16
      Top             =   2400
      Width           =   360
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Instructor"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   4080
      Width           =   720
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Room"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   13
      Top             =   3675
      Width           =   405
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Out"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   11
      Top             =   2835
      Width           =   645
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time In"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   9
      Top             =   2415
      Width           =   525
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Abr."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4260
      TabIndex        =   7
      Top             =   1275
      Width           =   900
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Title"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   1695
      Width           =   885
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CN"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   1320
      Width           =   210
   End
End
Attribute VB_Name = "frmSubjectEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mFormState As String

Dim curSubject As tSubject

Dim mShowAdd As Boolean
Dim mShowEdit As Boolean
Dim isON As Boolean

Public Function ShowAdd(Optional iSYID As Integer = 0, Optional bSem As Byte = 1, Optional bTerm As Byte = 1) As Boolean
    
    'set form state
    mFormState = "add"
    
    curSubject.FK_SYID = iSYID
    curSubject.Sem = bSem
    curSubject.Term = bTerm
    
    'show form
    Me.Show vbModal
    
    'return
    ShowAdd = mShowAdd
    
End Function

Public Function ShowEdit(ByVal sSubjectID As String) As Boolean
    
    'set form state
    mFormState = "edit"
    
    'set parameter
    curSubject.SubjectID = sSubjectID
    
    'show form
    Me.Show vbModal
    
    'return
    ShowEdit = mShowEdit
    
End Function

Private Sub cmbSem_Change()
    GenerateSubjectID
End Sub

Private Sub cmbSY_Change()
    GenerateSubjectID
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


Private Sub GenerateSubjectID()
    
    If modFunction.IsComboItemSelected(cmbSY) = False Then
        Exit Sub
    End If
    If modFunction.IsComboItemSelected(cmbSem) = False Then
        Exit Sub
    End If
    If GetTxtVal(txtCN.Text) < 1 Then
        Exit Sub
    End If
    
    'Subject ID = SYID + "-" + Sem + "-" + CN
    
    txtSubjectID.Text = GetKeyOnSplit(cmbSY.SelectedItem.Key, "//", 1) & _
                        "-" & Trim(cmbSem.Text) & _
                        "-" & Trim(txtCN.Text)
                        
End Sub

Private Function SaveEdit()

    Dim curSubject As tSubject
    Dim OldSubject As tSubject
    
    'check
    
    'CN
    If GetTxtVal(txtCN.Text) < 1 Then
        MsgBox "Please enter valid CN.", vbExclamation
        HLTxt txtCN
        Exit Function
    End If
    
    'SubjectAbr
    If Len(Trim(txtSubjectAbr.Text)) < 1 Then
        MsgBox "Please enter Subject Abr.", vbExclamation
        HLTxt txtSubjectAbr
        Exit Function
    End If
    
    'SubjectTitle
    If Len(Trim(txtSubjectTitle.Text)) < 1 Then
        MsgBox "Please enter Subject Title.", vbExclamation
        HLTxt txtSubjectTitle
        Exit Function
    End If

    'check sy
    If IsComboItemSelected(cmbSY) = False Then
        MsgBox "Please select valid School Year.", vbExclamation
        cmbSY.SetFocus
        Exit Function
    End If
    
    'check Sem
    If IsComboItemSelected(cmbSem) = False Then
        MsgBox "Please select valid Semester.", vbExclamation
        cmbSem.SetFocus
        Exit Function
    End If
    
    'check Term
    If IsComboItemSelected(cmbTerm) = False Then
        MsgBox "Please select valid Term.", vbExclamation
        cmbTerm.SetFocus
        Exit Function
    End If
    
    
    

    'set new Subject
    With curSubject
        .SubjectID = txtSubjectID.Text
        .CN = GetTxtVal(txtCN.Text)
        .SubjectTitle = txtSubjectTitle.Text
        .SubjectAbr = txtSubjectAbr.Text
        .FK_SYID = Val(GetKeyOnSplit(cmbSY.SelectedItem.Key, "//", 1))
        .PreReqCN = -1
        .TimeIn = GetTxtVal(txtTimeIn.Text)
        .TimeOut = GetTxtVal(txtTimeOut.Text)
        .Room = txtRoom.Text
        .SchedDay = txtSchedDay.Text
        .Sem = Val(cmbSem.Text)
        .Term = Val(cmbTerm.Text)
        .Instructor = txtInstructor.Text

    End With
    
    'try
    'add new Subject
    If EditSubject(curSubject) = True Then
        
        MsgBox "Subject entry was successfully updated.", vbInformation
        'set flag
        mShowEdit = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to update Subject entry.", vbExclamation
        'set flag
        mShowEdit = False
        
    End If
    
End Function

Private Function SaveAdd()

    Dim NewSubject As tSubject
    Dim OldSubject As tSubject
    
    'check
    
    'CN
    If GetTxtVal(txtCN.Text) < 1 Then
        MsgBox "Please enter valid CN.", vbExclamation
        HLTxt txtCN
        Exit Function
    End If
    
    'CN Duplication
    If GetSubjectByID(Val(GetKeyOnSplit(cmbSY.SelectedItem.Key, "//", 1)), OldSubject) = True Then
        MsgBox "The Subject with CN: '" & txtCN.Text & "' is already existed.", vbExclamation
        HLTxt txtCN
        Exit Function
    End If
    
    'SubjectAbr
    If Len(Trim(txtSubjectAbr.Text)) < 1 Then
        MsgBox "Please enter Subject Abr.", vbExclamation
        HLTxt txtSubjectAbr
        Exit Function
    End If

    
    'SubjectTitle
    If Len(Trim(txtSubjectTitle.Text)) < 1 Then
        MsgBox "Please enter Subject Title.", vbExclamation
        HLTxt txtSubjectTitle
        Exit Function
    End If

    'check sy
    If IsComboItemSelected(cmbSY) = False Then
        MsgBox "Please select valid School Year.", vbExclamation
        cmbSY.SetFocus
        Exit Function
    End If
    

    'check Sem
    If IsComboItemSelected(cmbSem) = False Then
        MsgBox "Please select valid Semester.", vbExclamation
        cmbSem.SetFocus
        Exit Function
    End If
    
    'check Term
    If IsComboItemSelected(cmbTerm) = False Then
        MsgBox "Please select valid Term.", vbExclamation
        cmbTerm.SetFocus
        Exit Function
    End If
    
    
    

    'set new Subject
    With NewSubject
        .SubjectID = txtSubjectID.Text
        .CN = GetTxtVal(txtCN.Text)
        .SubjectTitle = txtSubjectTitle.Text
        .SubjectAbr = txtSubjectAbr.Text
        .FK_SYID = Val(GetKeyOnSplit(cmbSY.SelectedItem.Key, "//", 1))
        .PreReqCN = -1
        .TimeIn = GetTxtVal(txtTimeIn.Text)
        .TimeOut = GetTxtVal(txtTimeOut.Text)
        .Room = txtRoom.Text
        .SchedDay = txtSchedDay.Text
        .Sem = Val(cmbSem.Text)
        .Term = Val(cmbTerm.Text)
        .Instructor = txtInstructor.Text

    End With
    
    'try
    'add new Subject
    If AddSubject(NewSubject) = True Then
        
        MsgBox "New Subject entry was added.", vbInformation
        'set flag
        mShowAdd = True
        'close form and return
        Unload Me
        
    Else
    
        MsgBox "Unable to add new Subject entry.", vbExclamation
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
            Me.Caption = "Add Subject"
            
            'refresh SY
            Refresh_SY curSubject.FK_SYID
            'refresh sem
            Refresh_Sem curSubject.Sem
            'refresh term
            Refresh_Term curSubject.Term
         
        Case "edit"
            
            'set caption
            Me.Caption = "Edit Subject"

            'get info
            If GetSubjectByID(curSubject.SubjectID, curSubject) = False Then
                'show failed
                MsgBox "Subject entry with Subject ID : '" & curSubject.CN & "' does not exist.", vbExclamation
                'close this form
                Unload Me
                Exit Sub
            End If
            
            'set form ui info
            With curSubject
                txtSubjectID.Text = .SubjectID
                txtCN.Text = .CN
                txtSubjectTitle.Text = .SubjectTitle
                txtSubjectAbr.Text = .SubjectAbr
                'refresh SY
                Refresh_SY .FK_SYID

                '.PreReqCN = -1
                'refresh sem
                Refresh_Sem .Sem
                'refresh term
                Refresh_Term .Term
                txtTimeIn.Text = .TimeIn
                txtTimeOut.Text = .TimeOut
                txtRoom.Text = .Room
                txtSchedDay.Text = .SchedDay
                txtInstructor.Text = .Instructor
            End With

            'disable some controls
            txtCN.Enabled = False

            
    End Select
    
End Sub


Private Sub Form_Load()
    isON = False
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


Private Sub Refresh_Term(Optional iTerm As Byte = 0)

    With cmbTerm
        .ComboItems.Clear
        .ComboItems.Add , , "1"
        .ComboItems.Add , , "2"
        .ComboItems.Add , , "3"
        
        If iTerm = 0 Then
            'select first item
            .ComboItems(1).Selected = True
        Else
            'select iTerm
            .ComboItems(iTerm).Selected = True
        End If
    End With
End Sub


Private Sub txtCN_Change()
    GenerateSubjectID
End Sub
