VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.Form frmAllSubject 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   0  'None
   Caption         =   "Subjects"
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
   Icon            =   "frmAllSubject.frx":0000
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
      Caption         =   "Manage Subject Entries"
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
      Top             =   4470
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
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   530
      TabIndex        =   2
      Top             =   390
      Width           =   7950
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
      Top             =   870
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
               Picture         =   "frmAllSubject.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin b8Controls4.LynxGrid3 listEntries 
         Height          =   2355
         Left            =   510
         TabIndex        =   5
         Top             =   390
         Width           =   4815
         _ExtentX        =   8493
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
         Top             =   0
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
End
Attribute VB_Name = "frmAllSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim bReadyToDisplay As Boolean
Dim bFormStarted As Boolean

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'pass keyinfo to mdiMain
    mdiMain.AFForm_KeyDown KeyCode, Shift
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
End Sub

Private Sub Form_Load()
    
    'set list columns
    With listEntries
    
        .Redraw = False
        
        .AddColumn "Subject ID", 100         '0
        .AddColumn "S.Y. ID", 200         '1
        .AddColumn "S.Y.", 90               '2
        .AddColumn "Sem", 60               '3
        .AddColumn "Term", 60               '4
        .AddColumn "CN", 70               '5
        .AddColumn "Subject Title", 120   '6
        .AddColumn "Abr.", 90      '7
        .AddColumn "Pre-Req. CN", 70      '8
        .AddColumn "Sched.", 60      '9
        .AddColumn "Room", 80      '10
        .AddColumn "Time", 90      '11
        .AddColumn "Instructor", 120      '12
        
        .RowHeightMin = 21
        .ImageList = ilList
        .Redraw = True
        .Refresh
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
    shpFooter.Move 2, 1, bgFooter.Width - 4, bgFooter.Height - 3
    
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
    
    Dim vRS As New ADODB.Recordset
    Dim sSQL As String
    Dim li As Long
    
    
    
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

        
    
    If GetTxtVal(mdiMain.b8DPSY.BoundData) > 0 Then
        sSQL_SY = " tblSubject.FK_SYID=" & mdiMain.b8DPSY.BoundData & " "
    Else
        sSQL_SY = ""
    End If

    If mdiMain.cmbSem.ListIndex > 0 Then
        sSQL_Sem = " tblSubject.Sem=" & mdiMain.cmbSem.Text & " "
    Else
        sSQL_Sem = ""
    End If
    
    If mdiMain.cmbTerm.ListIndex > 0 Then
        sSQL_Term = " tblSubject.Term=" & Val(mdiMain.cmbTerm.Text) & " "
    Else
        sSQL_Term = ""
    End If
    
    'combine
    sSQL_All = " WHERE  tblSubject.Sem>0" & IIf(Len(sSQL_SY) > 0, " AND " & sSQL_SY, "") '
    sSQL_All = sSQL_All & IIf(Len(sSQL_Sem) > 0, " AND " & sSQL_Sem, "")
    sSQL_All = sSQL_All & IIf(Len(sSQL_Term) > 0, " AND " & sSQL_Term, "")
   
    'set SQL Expression
    sSQL = "SELECT tblSubject.SubjectID, tblSubject.FK_SYID, tblSY.SYTitle, tblSubject.Sem, tblSubject.Term, tblSubject.CN, tblSubject.SubjectTitle, tblSubject.SubjectAbr, tblSubject.PreReqCN, tblSubject.TimeIn, tblSubject.TimeOut, tblSubject.Room, tblSubject.SchedDay, tblSubject.Instructor" & _
            " FROM tblSY INNER JOIN tblSubject ON tblSY.SYID = tblSubject.FK_SYID" & _
            sSQL_All & " ORDER BY tblSubject.FK_SYID, tblSubject.Sem, tblSubject.Term, tblSubject.SubjectTitle"


    If ConnectRS(PrimeData, vRS, sSQL) = False Then
        WriteErrorLog Me.Name, "LoadEntries", "Unable to connect Recordset. SQL Expression: '" & sSQL & "'"
        GoTo RAE
    End If
    
    If AnyRecordExisted(vRS) = False Then
        GoTo RAE
    End If
    
    'add entries to list
    vRS.MoveFirst
    While vRS.EOF = False
    
        With listEntries
            li = .AddItem(ReadField(vRS.Fields("SubjectID")))
            .ItemImage(li) = 1
            .CellText(li, 1) = ReadField(vRS.Fields("FK_SYID"))
            .CellText(li, 2) = ReadField(vRS.Fields("SYTitle"))
            .CellText(li, 3) = ReadField(vRS.Fields("Sem"))
            .CellText(li, 4) = ReadField(vRS.Fields("Term"))
            
            .CellFontBold(li, 5) = True
            .CellFontBold(li, 6) = True
            .CellFontBold(li, 7) = True
            .CellText(li, 5) = ReadField(vRS.Fields("CN"))
            .CellText(li, 6) = ReadField(vRS.Fields("SubjectTitle"))
            .CellText(li, 7) = ReadField(vRS.Fields("SubjectAbr"))
            
            .CellText(li, 8) = IIf(ReadField(vRS.Fields("PreReqCN")) > 0, ReadField(vRS.Fields("PreReqCN")), " ")
            .CellText(li, 9) = ReadField(vRS.Fields("SchedDay"))
            .CellText(li, 10) = ReadField(vRS.Fields("Room"))
            .CellText(li, 11) = ReadField(vRS.Fields("TimeIn")) & "-" & ReadField(vRS.Fields("TimeOut"))
            .CellText(li, 12) = ReadField(vRS.Fields("Instructor"))
        
        
        End With
        
        vRS.MoveNext
    Wend
    
RAE:
    Set vRS = Nothing
    listEntries.Redraw = True
    listEntries.Refresh
    'refresh rec sum
    RefreshRecSum
    'refresh recopt buttons
    mdiMain.ActivateChild Me
    'restore mouse pointer
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

    Form_CanAdd = True

End Function
Public Function Form_Add()
    
    If frmSubjectEntry.ShowAdd = True Then
        Form_Refresh
    End If
    
End Function

Public Function Form_CanEdit() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanEdit = True
    Else
        Form_CanEdit = False
    End If
End Function

Public Function Form_Edit()

    Dim sSubjectID As String
    
    If listEntries.RowCount < 1 Then
        Exit Function
    End If
    
    sSubjectID = listEntries.CellText(listEntries.Row, 0)
    
    If frmSubjectEntry.ShowEdit(sSubjectID) = True Then
        Form_Refresh
    End If
    
End Function

Public Function Form_CanDelete() As Boolean
    If listEntries.RowCount > 0 Then
        Form_CanDelete = True
    Else
        Form_CanDelete = False
    End If
End Function

Public Function Form_Delete()
    
    Dim sSubjectID As String
    
    If listEntries.RowCount < 1 Then
        Exit Function
    End If

    If MsgBox("Are you sure you want to delete this Subject entry named   '" & listEntries.CellText(listEntries.Row, 1) & "' ?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
        Exit Function
    End If
    
    'get ID
    sSubjectID = listEntries.CellText(listEntries.Row, 0)

    If DeleteSubject(sSubjectID) = True Then
        Form_Refresh
    Else
        WriteErrorLog Me.Name, "Form_Delete", "Failed on: 'DeleteSubject(sSubjectID) = True'"
    End If

End Function

Public Function Form_CanRefresh() As Boolean
    Form_CanRefresh = True
End Function

Public Function Form_Refresh()
    LoadEntries
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
                
        sFields(0) = "Subject ID" ' 100         '0
        sFields(1) = "S.Y. ID" ' 200         '1
        sFields(2) = "S.Y." ' 90               '2
        sFields(3) = "Sem" ' 60               '3
        sFields(4) = "Term" ' 60               '4
        sFields(5) = "CN" ' 70               '5
        sFields(6) = "Subject Title" ' 120   '6
        sFields(7) = "Abr." ' 90      '7
        sFields(8) = "Pre-Req. CN" ' 70      '8
        sFields(9) = "Sched." ' 60      '9
        sFields(10) = "Room" ' 80      '10
        sFields(12) = "Time" ' 90      '11
        sFields(13) = "Instructor" ' 120      '12
       
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


Public Sub Form_FilterChange()

    Me.Form_Refresh
   
End Sub




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


'----------------------------------------------------------
' >>> END Menu Procedures
'----------------------------------------------------------

Private Sub mnuDelete_Click()
    Form_Delete
End Sub

Private Sub mnuEdit_Click()
    Form_Edit
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
