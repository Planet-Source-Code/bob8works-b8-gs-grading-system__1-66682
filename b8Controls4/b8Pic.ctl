VERSION 5.00
Begin VB.UserControl b8Pic 
   BackColor       =   &H00F7F7F7&
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   177
   Begin VB.Image imgNoPic 
      Height          =   1545
      Left            =   2550
      Picture         =   "b8Pic.ctx":0000
      Top             =   2370
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image imgPic 
      Height          =   1815
      Left            =   330
      Top             =   300
      Width           =   1935
   End
End
Attribute VB_Name = "b8Pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_AutoFit = True
'Property Variables:
Dim m_AutoFit As Boolean

Dim m_PicPath As String
Dim m_PicLoaded As Boolean

Private Function GetHeight() As Integer
    GetHeight = UserControl.Height / Screen.TwipsPerPixelY
End Function

Private Function GetWidth() As Integer
    GetWidth = UserControl.Width / Screen.TwipsPerPixelX
End Function



Private Sub imgPic_DblClick()
    If Me.IsPicLoaded = True Then
        frmPicPrev.ShowForm imgPic
    End If
End Sub

Private Sub UserControl_Initialize()
    m_PicLoaded = False
End Sub

Private Sub UserControl_Resize()

    imgPic.Stretch = False

    If m_AutoFit = True And (imgPic.Width > GetWidth Or imgPic.Height > GetHeight) Then
        imgPic.Stretch = True
        imgPic.Move 0, 0, GetWidth, GetHeight
    Else
        imgPic.Move (GetWidth - imgPic.Width) / 2, (GetHeight - imgPic.Height) / 2
    End If

End Sub

Public Sub LoadPic(ByVal NewPicPath As String)
    
    On Error GoTo InvalidPic
    'default
    m_PicLoaded = False
    
    m_PicPath = NewPicPath
    'Load
    imgPic.Picture = LoadPicture(m_PicPath)
    Call UserControl_Resize
           
    m_PicLoaded = True
    
    Exit Sub
InvalidPic:
    imgPic.Stretch = False
    Set imgPic.Picture = imgNoPic.Picture
    Call UserControl_Resize
End Sub

Public Sub ClearPic()
    m_PicLoaded = False
    Set imgPic.Picture = Nothing
    imgPic.Refresh
End Sub

Public Property Get PicPath() As String
    PicPath = m_PicPath
End Property

Public Function IsPicLoaded() As Boolean
    IsPicLoaded = m_PicLoaded
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgPic,imgPic,-1,Stretch
Public Property Get Stretch() As Boolean
Attribute Stretch.VB_Description = "Returns/sets a value that determines whether a graphic resizes to fit the size of an Image control."
    Stretch = imgPic.Stretch
End Property

Public Property Let Stretch(ByVal New_Stretch As Boolean)
    imgPic.Stretch() = New_Stretch
    PropertyChanged "Stretch"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    imgPic.Stretch = PropBag.ReadProperty("Stretch", False)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    m_AutoFit = PropBag.ReadProperty("AutoFit", m_def_AutoFit)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Stretch", imgPic.Stretch, False)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("AutoFit", m_AutoFit, m_def_AutoFit)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoFit() As Boolean
    AutoFit = m_AutoFit
End Property

Public Property Let AutoFit(ByVal New_AutoFit As Boolean)
    m_AutoFit = New_AutoFit
    PropertyChanged "AutoFit"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AutoFit = m_def_AutoFit
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

