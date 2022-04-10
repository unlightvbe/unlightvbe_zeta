VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl 大人物形像 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   9405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9675
   ScaleHeight     =   9405
   ScaleWidth      =   9675
   Windowless      =   -1  'True
   Begin ImageX.aicAlphaImage aicimage1 
      Height          =   6180
      Left            =   0
      Top             =   0
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   10901
      Image           =   "uc大人物形像.ctx":0000
      Scaler          =   3
      Props           =   13
   End
End
Attribute VB_Name = "大人物形像"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_smallimage As String
Dim m_bighei As Integer
Dim m_bigwh As Integer
Public Property Get 大人物圖片() As String
   大人物圖片 = m_smallimage
End Property
Public Property Get 大人物圖片height() As Integer
   大人物圖片height = m_bighei
End Property
Public Property Get 大人物圖片width() As Integer
   大人物圖片width = m_bigwh
End Property
Public Property Let 大人物圖片(ByVal New_大人物圖片 As String)
   m_smallimage = New_大人物圖片
   PropertyChanged "大人物圖片"
   If Me.大人物圖片 <> "" Then
'       aicimage1.AutoSize = True
'       aicimage1.AutoRedraw = True
       aicimage1.LoadImage_FromFile Me.大人物圖片
       aicimage1.Top = 0
       aicimage1.Left = 0
    End If
    Me.大人物圖片height = aicimage1.Height
    Me.大人物圖片width = aicimage1.Width
End Property
Public Property Let 大人物圖片height(ByVal New_大人物圖片height As Integer)
   m_bighei = New_大人物圖片height
   PropertyChanged "大人物圖片height"
End Property
Public Property Let 大人物圖片width(ByVal New_大人物圖片width As Integer)
   m_bigwh = New_大人物圖片width
   PropertyChanged "大人物圖片width"
End Property

