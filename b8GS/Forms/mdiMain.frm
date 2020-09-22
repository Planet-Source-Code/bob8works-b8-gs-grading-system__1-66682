VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\..\b8Controls4\b8Controls4.vbp"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00F5F5F5&
   Caption         =   "GradingSystem"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer timeUpdateDate 
      Interval        =   1000
      Left            =   3690
      Top             =   1170
   End
   Begin b8Controls4.b8ClientWin b8CW 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   6150
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   661
   End
   Begin b8Controls4.b8SBCenter b8SBC 
      Align           =   3  'Align Left
      Height          =   5070
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   8943
      MinWidth        =   180
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   540
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Quick Launch              [ Ctrl + Q ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   2115
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         AutoContract    =   0   'False
         Begin MSComctlLib.ImageList ilQL 
            Left            =   1080
            Top             =   660
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
                  Picture         =   "mdiMain.frx":1CFA
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView listQL 
            Height          =   1785
            Left            =   30
            TabIndex        =   43
            Top             =   420
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   3149
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            _Version        =   393217
            Icons           =   "ilQL"
            SmallIcons      =   "ilQL"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDropMode     =   1
            NumItems        =   0
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   1
         Left            =   60
         TabIndex        =   3
         Top             =   840
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Search Item               [ Ctrl + S ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   2245
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         Begin VB.TextBox txtSearchWhat 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   6
            Top             =   690
            Width           =   3165
         End
         Begin VB.ComboBox cmbLookIn 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1470
            Width           =   3135
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   4
            Top             =   1950
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Search What:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   450
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Look In:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   7
            Top             =   1260
            Width           =   585
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   2
         Left            =   60
         TabIndex        =   11
         Top             =   1170
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Filter By Date             [ Ctrl + D ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   2865
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         Begin b8Controls4.b8DatePicker b8DateP 
            Height          =   2415
            Left            =   120
            TabIndex        =   12
            Top             =   420
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   4260
            BackColor       =   16777215
            MinDate         =   38968
            MaxDate         =   38968
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   3
         Left            =   60
         TabIndex        =   31
         Top             =   1470
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Filter                           [ Ctrl + F ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   2683
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         Begin VB.CheckBox chkAutoRefresh 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Auto Refresh"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   42
            Top             =   2400
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.ComboBox cmbYL 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1950
            Width           =   735
         End
         Begin VB.ComboBox cmbTerm 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1950
            Width           =   735
         End
         Begin VB.ComboBox cmbSem 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1950
            Width           =   735
         End
         Begin b8Controls4.b8DataPicker b8DPCourse 
            Height          =   345
            Left            =   120
            TabIndex        =   34
            Top             =   1320
            Width           =   3195
            _ExtentX        =   5636
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
         Begin b8Controls4.b8DataPicker b8DPSY 
            Height          =   345
            Left            =   120
            TabIndex        =   32
            Top             =   660
            Width           =   3195
            _ExtentX        =   5636
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Y.L."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1800
            TabIndex        =   41
            Top             =   1740
            Width           =   285
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Term:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   960
            TabIndex        =   39
            Top             =   1740
            Width           =   420
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Sem:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   1740
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Course:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1110
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&School Year:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   450
            Width           =   900
         End
      End
      Begin b8Controls4.b8SideTab b8ST 
         Height          =   345
         Index           =   4
         Left            =   60
         TabIndex        =   44
         Top             =   1800
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         Caption         =   "Student's Picture"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         MaxHeight       =   3910
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   9594695
         AutoContract    =   0   'False
         Begin b8Controls4.b8Pic b8PicStud 
            Height          =   3435
            Left            =   90
            TabIndex        =   45
            Top             =   390
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   6059
         End
         Begin VB.Label lblStudName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[Select student to view picture]"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   46
            Top             =   3870
            Width           =   2685
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today is "
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
         Height          =   165
         Left            =   150
         TabIndex        =   30
         Top             =   270
         Width           =   585
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   780
         TabIndex        =   29
         Top             =   255
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome,"
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
         Height          =   165
         Left            =   150
         TabIndex        =   28
         Top             =   60
         Width           =   600
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   27
         Top             =   45
         Width           =   180
      End
   End
   Begin VB.PictureBox bgHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   0
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   0
      Top             =   0
      Width           =   7515
      Begin VB.PictureBox bgRecOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   3600
         ScaleHeight     =   51
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1024
         TabIndex        =   17
         Top             =   330
         Width           =   15360
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   0
            Left            =   30
            TabIndex        =   18
            Top             =   60
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":1E0D
            BackColor       =   -2147483643
            Caption         =   "New"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":26E7
            BgColorDown     =   12632256
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   1
            Left            =   1050
            TabIndex        =   19
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":2FC1
            BackColor       =   -2147483643
            Caption         =   "Edit"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   4210752
            DisabledPicture =   "mdiMain.frx":389B
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   2
            Left            =   2070
            TabIndex        =   20
            Top             =   60
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":4175
            BackColor       =   -2147483643
            Caption         =   "Delete"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   4210752
            DisabledPicture =   "mdiMain.frx":4A4F
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   3
            Left            =   3270
            TabIndex        =   21
            Top             =   60
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":5329
            BackColor       =   -2147483643
            Caption         =   "Refresh"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":5C03
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8ToolButton b8RecOpt 
            Height          =   615
            Index           =   4
            Left            =   4560
            TabIndex        =   22
            Top             =   60
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1085
            Picture         =   "mdiMain.frx":64DD
            BackColor       =   -2147483643
            Caption         =   "Print"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   0
            DisabledPicture =   "mdiMain.frx":6DB7
            BorderColor     =   12632256
         End
         Begin b8Controls4.b8Line b8Line2 
            Height          =   30
            Left            =   0
            TabIndex        =   25
            Top             =   720
            Width           =   15720
            _ExtentX        =   27728
            _ExtentY        =   53
            BorderColor1    =   14737632
            BorderColor2    =   16777215
         End
      End
      Begin VB.PictureBox bgHeaderMenu 
         BackColor       =   &H00EDEBE9&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1024
         TabIndex        =   13
         Top             =   0
         Width           =   15360
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   14
            Top             =   15
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&System"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&System"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   15
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Records"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Records"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   2
            Left            =   1740
            TabIndex        =   23
            Top             =   15
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Tools"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Tools"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Menu b8Menus 
            Height          =   285
            Index           =   3
            Left            =   2370
            TabIndex        =   24
            Top             =   15
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Caption         =   "&Help"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Help"
            BorderColorNormal=   15592425
            BorderColorDown =   -2147483632
            BackColorNormal =   15592425
         End
         Begin b8Controls4.b8Line b8Line1 
            Height          =   30
            Left            =   0
            TabIndex        =   16
            Top             =   300
            Width           =   15720
            _ExtentX        =   27728
            _ExtentY        =   53
            BorderColor1    =   16119285
            BorderColor2    =   14737632
         End
      End
      Begin b8Controls4.b8Line b8LLogoB 
         Height          =   30
         Left            =   0
         TabIndex        =   26
         Top             =   1050
         Visible         =   0   'False
         Width           =   15720
         _ExtentX        =   27728
         _ExtentY        =   53
         BorderColor1    =   14737632
         BorderColor2    =   16777215
      End
      Begin b8Controls4.b8SBtop b8SBT 
         Height          =   945
         Left            =   0
         TabIndex        =   1
         Top             =   330
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1667
         MinWidth        =   180
         Begin VB.Image Image2 
            Height          =   210
            Left            =   690
            Picture         =   "mdiMain.frx":7691
            Top             =   300
            Width           =   1545
         End
         Begin VB.Image Image4 
            Height          =   135
            Left            =   720
            Picture         =   "mdiMain.frx":7A04
            Top             =   510
            Width           =   330
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   150
            Picture         =   "mdiMain.frx":7B26
            Top             =   150
            Width           =   480
         End
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&System"
      Visible         =   0   'False
      Begin VB.Menu mnuAddUser 
         Caption         =   "&Add New User"
      End
      Begin VB.Menu mnuManageUser 
         Caption         =   "&Manage Users"
      End
      Begin VB.Menu mnuS01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Records"
      Visible         =   0   'False
      Begin VB.Menu mnuSY 
         Caption         =   "S.Y. Entries"
         Begin VB.Menu mnuNewSY 
            Caption         =   "&New S.Y. Entry"
         End
         Begin VB.Menu mnuManageSY 
            Caption         =   "Manage S.Y. Entries"
         End
      End
      Begin VB.Menu mnuCourseEntries 
         Caption         =   "Course Entries"
         Begin VB.Menu mnuNewCourse 
            Caption         =   "&New Course Entry"
         End
         Begin VB.Menu mnuManageCourseEntries 
            Caption         =   "&Manage Course Entries"
         End
      End
      Begin VB.Menu mnuSubjectEntries 
         Caption         =   "&Subject Entries"
         Begin VB.Menu mnuNewSubject 
            Caption         =   "&New Subject Entry"
         End
         Begin VB.Menu mnuManageSubjectS 
            Caption         =   "&Manage Subject Entries"
         End
         Begin VB.Menu mnuS07 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuS10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStudentEntries 
         Caption         =   "&Student Records"
         Begin VB.Menu mnuNewStudent 
            Caption         =   "&New Student Entry"
         End
         Begin VB.Menu mnuManageStudent 
            Caption         =   "&Manage Student Entries"
         End
         Begin VB.Menu mnuS11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuManageEnrolledStudents 
            Caption         =   "&Manage Enrolled Students"
         End
         Begin VB.Menu mnuManageStudentGrades 
            Caption         =   "&Manage Student's Grades"
         End
         Begin VB.Menu mnuManageSubjectStudent 
            Caption         =   "&Manage Ranking List"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuPreferences 
         Caption         =   "&Preferences"
      End
      Begin VB.Menu mnuDatabaseUtilities 
         Caption         =   "&Database Utilities"
         Begin VB.Menu mnuBackupDatabase 
            Caption         =   "&Backup Database"
         End
         Begin VB.Menu mnuDatabaseRestore 
            Caption         =   "Database Restore"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuS03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutApp 
         Caption         =   "About &SIM"
      End
      Begin VB.Menu mnuAboutAuthor 
         Caption         =   "&About The Author"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const m_TabShowQuickLaunch = 0
Private Const m_TabSearch = 1
Private Const m_TabFilterDate = 2
Private Const m_TabFilter = 3
Private Const m_TabStudDetail = 4

'Flag for User log
Public bUserLoggedOn As Boolean


Public Function ShowForm()
    
    'default
    bUserLoggedOn = False
    
    'show form
    Me.WindowState = vbMaximized
    Me.Show
    DoEvents
    
    'show weclome
    frmWelcome.ShowForm
    
    'unload splash
    frmSplash.UnloadSplash
    
BeginLogin:
    If AnyUserExist = False Then
        If frmUserEntry.ShowAddAdmin = False Then
            Unload Me
        Else
            GoTo BeginLogin
        End If
    Else
        If frmLogin.ShowForm = False Then
            Unload Me
            Exit Function
        End If
    End If
    
    'set log flag
    bUserLoggedOn = True
    
      
    'set UI
    'current user info
    lblCurrentUser.Caption = CurrentUser.UserID
    
    'set date
    timeUpdateDate_Timer
    
    
    'if not admin, disable user menu
    If LCase(Trim(CurrentUser.UserID)) <> "administrator" Then
        mnuAddUser.Enabled = False
        mnuManageUser.Enabled = False
    Else
        mnuAddUser.Enabled = True
        mnuManageUser.Enabled = True
    End If
    


    'expand quick launch tab
    b8ST(m_TabShowQuickLaunch).Expanded = True
    
    
    frmAllStudent.ShowForm
    

End Function



Private Sub b8CW_CloseClick(ByVal sFormName As String, ByVal Index As Integer)
    'close form
    Dim frm As Form
    
    On Error GoTo RAE
    
    For Each frm In Forms
        If LCase(Trim(frm.Name)) = LCase(Trim(sFormName)) Then
            Unload frm
            Exit For
        End If
    Next
    
RAE:
    Set frm = Nothing
End Sub

'Control Procedures
'-----------------------------------------------------------
Private Sub b8CW_FormTabClick(ByVal sFormName As String, ByVal Index As Integer)
    modFuncChild.ActivateMDIChildForm sFormName
End Sub

Private Sub b8DateP_Change()
    Call Form_DateChange
End Sub





Private Sub b8DPCourse_Change()
    Form_FilterChange
End Sub

Private Sub b8DPSY_Change()
    Form_FilterChange
End Sub

Private Sub b8RecOpt_Click(Index As Integer)
    Select Case Index
        Case 0 'add
            Form_Add
        Case 1 'edit
            Form_Edit
        Case 2 'delete
            Form_Delete
        Case 3 'refresh
            Form_Refresh
        Case 4 'print
            Form_Print
    End Select
End Sub



Private Sub b8SBC_BeforeResize(ByVal NewWidth As Integer)
    ResizeFb8SBC NewWidth
End Sub

Private Sub ResizeFb8SBC(ByVal NewWidth As Integer)
    
    'resize top side bar
    b8SBT.Width = NewWidth / Screen.TwipsPerPixelX

    'resize quick tabs
    Dim i As Integer
    For i = 0 To b8ST.UBound
        b8ST(i).Left = 60
        b8ST(i).Width = NewWidth - 120
    Next
    
    'resize window tab
    If b8SBC.Visible = True Then
        b8CW.SBWidth = NewWidth / Screen.TwipsPerPixelX
    Else
        b8CW.SBWidth = 0
    End If
    
    'call mdi resize to resize all opened mdi childs
    MDIForm_Resize
    
End Sub

Private Sub b8SBC_Resize()
    ResizeFb8SBC b8SBC.Width
End Sub

Private Sub b8SBT_Resize()
    b8SBC.Width = b8SBT.Width * Screen.TwipsPerPixelX
End Sub

Private Sub b8SBT_SizeChange(ByVal newSizeState As b8Controls4.eSizeState)
    
    If newSizeState = ssContracted Then
        b8CW.SBWidth = b8SBC.Width / Screen.TwipsPerPixelX
        b8SBC.Visible = True
        b8LLogoB.Visible = False
    Else
        b8CW.SBWidth = 0
        b8SBC.Visible = False
        b8LLogoB.Visible = True
    End If
    
    'call mdi resize to resize all opened child forms
    Call MDIForm_Resize
    
End Sub

Private Sub b8ST_BeforeExpand(Index As Integer)

    'resize contained controlsbeofre expanding
    Select Case Index
        Case m_TabShowQuickLaunch
            listQL.Move 90, listQL.Top, b8ST(Index).Width - 150

        Case m_TabSearch 'search
            'resize
            txtSearchWhat.Move 90, txtSearchWhat.Top, b8ST(Index).Width - 150
            cmbLookIn.Move 90, cmbLookIn.Top, txtSearchWhat.Width
            cmdSearch.Move b8ST(Index).Width - cmdSearch.Width - 90
        Case m_TabFilterDate 'filter date
            b8DateP.Move 90, b8DateP.Top, b8ST(Index).Width - 150
        Case m_TabFilter 'filter
            b8DPSY.Move 90, b8DPSY.Top, b8ST(Index).Width - 150
            b8DPCourse.Move 90, b8DPCourse.Top, b8ST(Index).Width - 150
        
        Case m_TabStudDetail
            If Form_CanShowStudentDetail = False Then
                MsgBox "There is no Student's Picture to display.", vbInformation
                b8ST(Index).Expanded = False
            Else
                b8PicStud.Move 90, b8PicStud.Top, b8ST(Index).Width - 150
            End If
    End Select

End Sub

Private Sub b8ST_CompleteExpand(Index As Integer)
    Dim i As Integer
    
    For i = 0 To b8ST.UBound
        If Index <> i Then
            If b8ST(i).AutoContract = True Then
                b8ST(i).Expanded = False
            End If
        End If
    Next

End Sub

Private Sub b8ST_Resize(Index As Integer)
    
    Dim i As Integer
    
    For i = 1 To b8ST.UBound
        b8ST(i).Move b8ST(i).Left, (b8ST(i - 1).Top + b8ST(i - 1).Height) - 15
    Next
    
    If b8ST(Index).Expanded = True Then
        Select Case Index
            Case m_TabShowQuickLaunch
                listQL.Move 90, listQL.Top, b8ST(Index).Width - 150

            Case m_TabSearch 'search
                'resize
                txtSearchWhat.Move 90, txtSearchWhat.Top, b8ST(Index).Width - 150
                cmbLookIn.Move 90, cmbLookIn.Top, txtSearchWhat.Width
                cmdSearch.Move b8ST(Index).Width - cmdSearch.Width - 90
            Case m_TabFilterDate 'filter date
                b8DateP.Move 90, b8DateP.Top, b8ST(Index).Width - 150
            Case m_TabFilter
                b8DPSY.Move 90, b8DPSY.Top, b8ST(Index).Width - 150
                b8DPCourse.Move 90, b8DPCourse.Top, b8ST(Index).Width - 150
            Case m_TabStudDetail
                b8PicStud.Move 90, b8PicStud.Top, b8ST(Index).Width - 150
        End Select
    End If

End Sub





Private Sub cmbLookIn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdSearch_Click
    End If
End Sub



Private Sub cmbSem_Change()
    Form_FilterChange
End Sub

Private Sub cmbSem_Click()
    Form_FilterChange
End Sub

Private Sub cmbTerm_Change()
    Form_FilterChange
End Sub

Private Sub cmbTerm_Click()
    Form_FilterChange
End Sub

Private Sub cmbYL_Change()
    Form_FilterChange
End Sub

Private Sub cmbYL_Click()
    Form_FilterChange
End Sub



Private Sub listQL_DblClick()

    Dim selItemKey As String
    
    On Error GoTo RAE
    
    selItemKey = listQL.SelectedItem.Key
    
    Select Case selItemKey
        Case "sy" 'Manage S.Y. Entries"
            frmAllSY.ShowForm
        Case "course" 'Manage CourseEntries"
            frmAllCourse.ShowForm
        Case "subject" 'Manage Subjects"
            frmAllSubject.ShowForm
        Case "student" 'Manage Students"
            frmAllStudent.ShowForm
        Case "studrec" 'Manage Student's Record"
            frmStudentRecord.ShowForm
        Case "rank" 'Manage Ranking List"
            frmSubjectStudent.ShowForm
    End Select
    
RAE:

End Sub

Private Sub MDIForm_Load()
    
    
    'set menus
    Set b8Menus(0).Menu = Me.mnuSystem
    Set b8Menus(1).Menu = Me.mnuRecords
    Set b8Menus(2).Menu = Me.mnuTools
    Set b8Menus(3).Menu = Me.mnuHelp
    
    
    With b8DPSY
        Set .DropDBCon = PrimeData
        .SQLFields = "tblSY.SYID, tblSY.SYTitle"
        .SQLTable = "tblSY"
        .SQLWhereFields = " CStr(tblSY.SYID) & ' ' & tblSY.SYTitle "
        .SQLOrderBy = "tblSY.SYID"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 1
        .AddColumn "ID", 80
        .AddColumn "Title", 300
        
    End With
    
    
    With b8DPCourse
        Set .DropDBCon = PrimeData
        .SQLFields = "tblCourse.CourseID, tblCourse.Title, tblCourse.Description"
        .SQLTable = "tblCourse"
        .SQLWhereFields = " CStr(tblCourse.CourseID) & ' ' & tblCourse.Title & ' ' & tblCourse.Description "
        .SQLOrderBy = "tblCourse.Title"
        
        .BoundFieldIndex = 0 'Bound Index
        .DisplayFieldIndex = 1
        .AddColumn "ID", 60
        .AddColumn "Course", 120
        .AddColumn "Description", 200
        
    End With
    
    With cmbSem
        .AddItem "All"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .ListIndex = 0
    End With
    
    With cmbTerm
        .AddItem "All"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .ListIndex = 0
    End With
    
    With cmbYL
        .AddItem "All"
        .AddItem "I"
        .AddItem "II"
        .AddItem "III"
        .AddItem "IV"
        .AddItem "V"
        .AddItem "VI"
        .ListIndex = 0
    End With


    'add Quick Launch Items
    With listQL.ListItems
        .Add , "sy", "Manage S.Y. Entries", 1, 1
        .Add , "course", "Manage Course Entries", 1, 1
        .Add , "subject", "Manage Subjects", 1, 1
        .Add , "student", "Manage Students", 1, 1
        .Add , "studrec", "Manage Student's Record", 1, 1
        .Add , "rank", "Manage Ranking List", 1, 1
    End With
    
    
    'set filter default values
    b8DPSY.BoundData = CurrentSY.SYID
    b8DPSY.DisplayData = CurrentSY.SYTitle
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim FSO As New FileSystemObject
    Dim f As TextStream
    Dim s As String
    
    On Error GoTo errh
    
    If Command <> "be" Then
        MsgBox "Thank you for downloading b8 Grading System." & vbNewLine & vbNewLine & _
                "A single vote from you makes me inspire to share more codes on PSCODE.com. Please don't forget to vote.", vbExclamation
        
        'open web browser
        Set f = FSO.OpenTextFile(App.Path & "\URL.txt", ForReading)
        modFunction.OpenURL f.ReadLine, Me.hWnd
       
    End If
    
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
    Next
    
errh:
    Set frm = Nothing
    Set FSO = Nothing
    Set f = Nothing
    
End Sub

Private Sub mnuAboutApp_Click()
    frmSplash.ShowForm
End Sub

Private Sub mnuAboutAuthor_Click()
    frmAboutAuthor.ShowForm
End Sub

Private Sub mnuAddUser_Click()
    frmUserEntry.ShowAdd
End Sub

Private Sub mnuBackupDatabase_Click()
    frmDBBackup.ShowForm
End Sub

Private Sub mnuContents_Click()
    modFunction.OpenURL App.Path & "\help\index.htm", Me.hWnd
End Sub

Private Sub mnuDatabaseRestore_Click()
    frmRestore.ShowForm
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuLogOff_Click()
    Me.ShowForm
End Sub

Private Sub mnuManageCourseEntries_Click()
    frmAllCourse.ShowForm
End Sub

Private Sub mnuManageEnrolledStudents_Click()
    frmAllEnrollment.ShowForm
End Sub

Private Sub mnuManageStudent_Click()
    frmAllStudent.ShowForm
End Sub

Private Sub mnuManageStudentGrades_Click()
    frmStudentRecord.ShowForm
End Sub

Private Sub mnuManageSubjects_Click()
    frmAllSubject.ShowForm
End Sub

Private Sub mnuManageSubjectStudent_Click()
    frmSubjectStudent.ShowForm
End Sub

Private Sub mnuManageSY_Click()
    frmAllSY.ShowForm
End Sub

Private Sub mnuManageUser_Click()
    frmAllUser.ShowForm
End Sub

Private Sub mnuNewCourse_Click()
    frmCourseEntry.ShowAdd
End Sub

Private Sub mnuNewStudent_Click()
    frmStudentEntry.ShowAdd
End Sub

Private Sub mnuNewSubject_Click()
    frmSubjectEntry.ShowAdd
End Sub

Private Sub mnuNewSY_Click()
    frmSYEntry.ShowAdd
End Sub

Private Sub timeUpdateDate_Timer()
    lblDate.Caption = FormatDateTime(Now, vbGeneralDate)
End Sub

Private Sub txtSearchWhat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdSearch_Click
    End If
End Sub

Private Sub cmdSearch_Click()
    Form_Search
End Sub

'-----------------------------------------------------------
' end Control Procedures


' MDI Form procedures
'-----------------------------------------------------------
Private Sub MDIForm_Resize()
        
    Dim frm As Form
    
    
    On Error Resume Next
    
    'resize header menus bg
    'bgHeaderMenu.Left = b8SBC.Width / Screen.TwipsPerPixelX
    
    'resize bg Record Opt
    bgRecOpt.Move b8SBC.Width / Screen.TwipsPerPixelX
    
    
    'resize childs
    If GetActiveChildCount > 0 Then
        For Each frm In Forms
        If frm.Name <> Me.Name Then
            If frm.MDIChild = True Then
                If frm.Name = Me.ActiveForm.Name Then
                    ResizeMdiChildForm frm
                Else
                    frm.Visible = False
                End If
            End If
        End If
        
        Next
        
    End If
    
    Set frm = Nothing
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    'close database
    CloseDB modMain.PrimeData
End Sub
'Get Opened MDI Child Forms Count
Public Function GetActiveChildCount() As Integer
    
    Dim frm As Form
    Dim iCount As Integer
    
    iCount = 0
    
    For Each frm In Forms
        If frm.Name <> Me.Name Then
            If frm.MDIChild = True Then
                iCount = iCount + 1
            End If
        End If
    Next
    
    GetActiveChildCount = iCount
    Set frm = Nothing
    
End Function

'-----------------------------------------------------------
' >> End MDI Form procedures
'------------------------------------------------------------



'------------------------------------------------------------
' Parent To Child procedures
'------------------------------------------------------------

Public Sub AddChild(ByRef CFrm As Form, Optional CloseButton As Boolean = True)

    'load form
    modFuncChild.LoadForm CFrm, CloseButton
    
End Sub



Public Sub ActivateChild(ByRef CFrm As Form)

    'activate form
    Me.b8CW.SetActiveWindow CFrm.Name
    
    'refresh record operation buttons
    Form_CanAdd
    Form_CanEdit
    Form_CanDelete
    Form_CanRefresh
    Form_CanPrint
    Form_CanSearch
    Form_SetSearch
    Form_CanShowStudentDetail

    
End Sub

Public Sub RemoveChild(ByVal sFormName As String)
    
    'remove form
    Me.b8CW.RemoveChildWindow sFormName
    
End Sub



'Record Operation

Public Function Form_CanAdd() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanAdd
    
    b8RecOpt(0).Enabled = bReturn

    Form_CanAdd = bReturn
    
    Err.Clear
    
End Function
Public Function Form_Add()
    
    If Form_CanAdd Then
        Me.ActiveForm.Form_Add
    End If

End Function


Public Function Form_CanEdit() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanEdit
    
    b8RecOpt(1).Enabled = bReturn

    Form_CanEdit = bReturn
    
    Err.Clear
    
End Function
Public Function Form_Edit()
    
    If Form_CanEdit Then
        Me.ActiveForm.Form_Edit
    End If

End Function


Public Function Form_CanDelete() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanDelete
    
    b8RecOpt(2).Enabled = bReturn

    Form_CanDelete = bReturn
    
    Err.Clear
    
End Function


Public Function Form_Delete()
    
    If Form_CanDelete Then
        Me.ActiveForm.Form_Delete
    End If

End Function


Public Function Form_CanRefresh() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanRefresh
    
    b8RecOpt(3).Enabled = bReturn

    Form_CanRefresh = bReturn
    
    Err.Clear
    
End Function


Public Function Form_Refresh()
    
    If Form_CanRefresh Then
        Me.ActiveForm.Form_Refresh
    End If

End Function



Public Function Form_CanPrint() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanPrint
    
    b8RecOpt(4).Enabled = bReturn

    Form_CanPrint = bReturn
    
    Err.Clear
    
End Function


Public Function Form_Print()
    
    If Form_CanPrint Then
        Me.ActiveForm.Form_Print
    End If

End Function


Public Function Form_CanSearch() As Boolean

    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanSearch
    
    Form_CanSearch = bReturn
    
    Err.Clear
    
End Function



Public Function Form_ShowQuickLaunch()

    'expand side bar
    If b8SBT.SizeState <> ssContracted Then
        b8SBT.SizeState = ssContracted
    End If

    'expand search tab
    If b8ST(m_TabShowQuickLaunch).Expanded = False Then
        b8ST(m_TabShowQuickLaunch).Expanded = True
    End If
    
    On Error Resume Next
    b8ST(m_TabShowQuickLaunch).SetFocus
    'HLTxt txtSearchWhat
    Err.Clear
    
End Function

Public Function Form_ShowSearch()

    'expand side bar
    If b8SBT.SizeState <> ssContracted Then
        b8SBT.SizeState = ssContracted
    End If

    'expand search tab
    If b8ST(m_TabSearch).Expanded = False Then
        b8ST(m_TabSearch).Expanded = True
    End If
    
    On Error Resume Next
    b8ST(m_TabSearch).SetFocus
    HLTxt txtSearchWhat
    Err.Clear
    
End Function


Public Function Form_ShowDateFilter()

    'expand side bar
    If b8SBT.SizeState <> ssContracted Then
        b8SBT.SizeState = ssContracted
    End If

    'expand search tab
    If b8ST(m_TabFilterDate).Expanded = False Then
        b8ST(m_TabFilterDate).Expanded = True
    End If
    
    On Error Resume Next
    b8ST(m_TabFilterDate).SetFocus
    b8DateP.SetFocus
    Err.Clear
    
End Function


Public Function Form_SetSearch()
    Dim bReturn As Boolean
    Dim sFields() As String
    Dim i  As Integer
    
    'clear
    txtSearchWhat.Text = ""
    cmbLookIn.Clear
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_SetSearch(sFields)

    txtSearchWhat.Enabled = bReturn
    cmbLookIn.Enabled = bReturn
    cmdSearch.Enabled = bReturn
    
    If bReturn = True Then
        cmbLookIn.AddItem "All Fields"
        cmbLookIn.ListIndex = 0
        If UBound(sFields) >= 0 Then
            For i = 0 To UBound(sFields)
                cmbLookIn.AddItem sFields(i)
            Next
        End If
    Else
        'contract search tab if it was expanded
        If b8ST(m_TabSearch).Expanded = True Then
            b8ST(m_TabSearch).Expanded = False
        End If
        
    End If
    
    Form_SetSearch = bReturn
    
    Err.Clear
End Function


Public Function Form_Search()
        
    Dim bResult As Boolean
    
    'default
    bResult = False
    
    
    On Error GoTo errh
    
    If txtSearchWhat.Text = "Enter text here" Then
        txtSearchWhat.Text = ""
    End If
    
    If Len(Trim(txtSearchWhat.Text)) <= 0 Then
        MsgBox "Please enter text to search.", vbExclamation
        txtSearchWhat.Text = "Enter text here"
        HLTxt txtSearchWhat
        GoTo errh
    End If
    
    If Len(Trim(cmbLookIn.Text)) <= 0 Then
        MsgBox "Please enter valid field.", vbExclamation
        cmbLookIn.SetFocus
        GoTo errh
    End If
    

    bResult = Me.ActiveForm.Form_Search(Trim(txtSearchWhat.Text), Trim(cmbLookIn.Text))

    If bResult = False Then
        MsgBox "Cannot find '" & txtSearchWhat.Text & "'", vbExclamation
        HLTxt txtSearchWhat
    End If
    
errh:
    Err.Clear
    
End Function

Public Sub Form_DateChange()

    On Error GoTo errh
    Me.ActiveForm.Form_DateChange
errh:

End Sub

Public Sub Form_FilterChange(Optional ByVal bForce As Boolean)

    If bForce = False Then
        If chkAutoRefresh.Value <> vbChecked Then
            Exit Sub
        End If
    End If
    
    
    On Error GoTo errh
    Me.ActiveForm.Form_FilterChange
errh:

End Sub



Public Function Form_StartBussy()
    Me.MousePointer = vbHourglass
End Function

Public Function Form_EndBussy()
    Me.MousePointer = vbDefault
End Function

Public Sub AFForm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 83 And Shift = 4 Then
        b8Menus(0).ShowPopUp
    ElseIf KeyCode = 82 And Shift = 4 Then
        b8Menus(1).ShowPopUp
    ElseIf KeyCode = 77 And Shift = 4 Then
        b8Menus(2).ShowPopUp
    ElseIf KeyCode = 84 And Shift = 4 Then
        b8Menus(3).ShowPopUp
    ElseIf KeyCode = 72 And Shift = 4 Then
        b8Menus(4).ShowPopUp
        
    ElseIf KeyCode = 81 And Shift = 2 Then
        'Ctrl + Q
        Me.Form_ShowQuickLaunch
    ElseIf KeyCode = 68 And Shift = 2 Then
        'Ctrl + D
        Me.Form_ShowDateFilter
    End If
    
    'MsgBox KeyCode & " - " & Shift
End Sub




Public Function Form_CanShowStudentDetail() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanShowStudentDetail
            
    If bReturn = False Then
        b8ST(m_TabStudDetail).Expanded = False
    End If
    
    Form_CanShowStudentDetail = bReturn
    
    Err.Clear
    
End Function


Public Function Form_ShowStudentDetail()
    
    Dim lStudentID As Long
    Dim vStudent As tStudent
    
    
    On Error GoTo RAE
    If Form_CanShowStudentDetail Then
        
        'show detail
        lStudentID = Me.ActiveForm.Form_GetStudentID
    
        b8PicStud.LoadPic App.Path & "/studpics/" & lStudentID & ".img"
        'load full name
        If GetStudentByID(lStudentID, vStudent) = True Then
            lblStudName.Caption = vStudent.FirstName & " " & vStudent.MiddleName & " " & vStudent.LastName
        Else
            lblStudName.Caption = "[Select student to view picture]"
        End If
    End If
    
    Exit Function
RAE:

End Function
'------------------------------------------------------------
' >>> Parent To Child procedures


'Member variables property
Public Property Get TabSearchIndex() As Integer
    TabSearchIndex = m_TabSearch
End Property


