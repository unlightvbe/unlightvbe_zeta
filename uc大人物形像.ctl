VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Imagex.ocx"
Begin VB.UserControl �j�H���ι� 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BackStyle       =   0  '�z��
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
      Image           =   "uc�j�H���ι�.ctx":0000
      Scaler          =   3
      Props           =   13
   End
End
Attribute VB_Name = "�j�H���ι�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_smallimage As String
Dim m_bighei As Integer
Dim m_bigwh As Integer
Public Property Get �j�H���Ϥ�() As String
   �j�H���Ϥ� = m_smallimage
End Property
Public Property Get �j�H���Ϥ�height() As Integer
   �j�H���Ϥ�height = m_bighei
End Property
Public Property Get �j�H���Ϥ�width() As Integer
   �j�H���Ϥ�width = m_bigwh
End Property
Public Property Let �j�H���Ϥ�(ByVal New_�j�H���Ϥ� As String)
   m_smallimage = New_�j�H���Ϥ�
   PropertyChanged "�j�H���Ϥ�"
   If Me.�j�H���Ϥ� <> "" Then
'       aicimage1.AutoSize = True
'       aicimage1.AutoRedraw = True
       aicimage1.LoadImage_FromFile Me.�j�H���Ϥ�
       aicimage1.Top = 0
       aicimage1.Left = 0
    End If
    Me.�j�H���Ϥ�height = aicimage1.Height
    Me.�j�H���Ϥ�width = aicimage1.Width
End Property
Public Property Let �j�H���Ϥ�height(ByVal New_�j�H���Ϥ�height As Integer)
   m_bighei = New_�j�H���Ϥ�height
   PropertyChanged "�j�H���Ϥ�height"
End Property
Public Property Let �j�H���Ϥ�width(ByVal New_�j�H���Ϥ�width As Integer)
   m_bigwh = New_�j�H���Ϥ�width
   PropertyChanged "�j�H���Ϥ�width"
End Property

