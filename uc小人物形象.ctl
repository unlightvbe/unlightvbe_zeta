VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl 小人物形象 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2880
   ClipBehavior    =   0  '無
   ScaleHeight     =   5535
   ScaleWidth      =   2880
   Windowless      =   -1  'True
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   4800
   End
   Begin ImageX.aicAlphaImage image1 
      Height          =   3255
      Left            =   120
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   5741
      Image           =   "uc小人物形象.ctx":0000
      Scaler          =   3
   End
   Begin ImageX.aicAlphaImage image2 
      Height          =   855
      Left            =   120
      Top             =   3240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      Image           =   "uc小人物形象.ctx":0018
      Scaler          =   3
   End
End
Attribute VB_Name = "小人物形象"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_totwidth As Integer
Dim m_totheight As Integer
Dim m_smalldowntop As Integer
Dim m_smallimage As String
Dim m_smallimagedown As String
Dim m_smalldownleft As Integer
Dim m_smallhei As Integer
Dim m_smallwh As Integer
Dim m_smalldeath As Boolean

Public Property Get 小人物影子Left() As Integer
   小人物影子Left = m_smalldownleft
End Property
Public Property Get 小人物影子top差() As Integer
   小人物影子top差 = m_smalldowntop
End Property
Public Property Get 小人物圖片() As String
   小人物圖片 = m_smallimage
End Property
Public Property Get 小人物圖片height() As Integer
   小人物圖片height = m_smallhei
End Property
Public Property Get 小人物圖片width() As Integer
   小人物圖片width = m_smallwh
End Property
Public Property Get 小人物影子圖片() As String
   小人物影子圖片 = m_smallimagedown
End Property

Public Property Let 小人物影子Left(ByVal New_小人物影子Left As Integer)
   m_smalldownleft = New_小人物影子Left
   PropertyChanged "小人物影子Left"
   image2.Left = Me.小人物影子Left
End Property
Public Property Let 小人物影子top差(ByVal New_小人物影子top差 As Integer)
   m_smalldowntop = New_小人物影子top差
   PropertyChanged "小人物影子top差"
   image2.Top = Image1.Height + Me.小人物影子top差
End Property
Public Property Let 小人物圖片(ByVal New_小人物圖片 As String)
   m_smallimage = New_小人物圖片
   PropertyChanged "小人物圖片"
   If Me.小人物圖片 <> "" Then
       Image1.AutoRedraw = True
       Image1.AutoSize = True
       Image1.LoadImage_FromFile Me.小人物圖片
       Image1.Left = 0
       Image1.Top = 0
       Me.小人物圖片height = Image1.Height
       Me.小人物圖片width = Image1.Width
       Image1.Opacity = 100
       Me.小人物消失 = False
   End If
End Property
Public Property Let 小人物影子圖片(ByVal New_小人物影子圖片 As String)
   m_smallimagedown = New_小人物影子圖片
   PropertyChanged "小人物影子圖片"
   If Me.小人物影子圖片 <> "" Then
       image2.AutoRedraw = True
       image2.AutoSize = True
       image2.LoadImage_FromFile Me.小人物影子圖片
       image2.Left = 0
       image2.Top = Image1.Height
       image2.Opacity = 100
   End If
End Property
Public Property Let 小人物圖片height(ByVal New_小人物圖片height As Integer)
   m_smallhei = New_小人物圖片height
   PropertyChanged "小人物圖片height"
End Property
Public Property Let 小人物圖片width(ByVal New_小人物圖片width As Integer)
   m_smallwh = New_小人物圖片width
   PropertyChanged "小人物圖片width"
End Property
Public Property Get 小人物消失() As Boolean
   小人物消失 = m_smalldeath
End Property
Public Property Let 小人物消失(ByVal New_小人物消失 As Boolean)
   m_smalldeath = New_小人物消失
   PropertyChanged "小人物消失"
   '=====================
   If Me.小人物消失 = True Then
       t1.Enabled = True
   End If
End Property

Private Sub t1_Timer()
If Image1.Opacity <> 0 Then
    Image1.Opacity = Val(Image1.Opacity) - 1
End If
If image2.Opacity <> 0 Then
    image2.Opacity = Val(image2.Opacity) - 1
End If
If Image1.Opacity = 0 And image2.Opacity = 0 Then
    t1.Enabled = False
    Me.小人物消失 = False
End If
End Sub

