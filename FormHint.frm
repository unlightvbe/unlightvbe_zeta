VERSION 5.00
Begin VB.Form FormHint 
   BorderStyle     =   1  '單線固定
   Caption         =   "UnlightVBE系統提示"
   ClientHeight    =   2805
   ClientLeft      =   3360
   ClientTop       =   4395
   ClientWidth     =   9120
   Icon            =   "FormHint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   9120
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   360
      Picture         =   "FormHint.frx":0CCA
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label bnet 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "稍等一下，大小姐。您還沒有完成設定歐。"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '不透明
      Height          =   1845
      Left            =   0
      Top             =   0
      Width           =   11295
   End
   Begin VB.Image bne 
      Height          =   615
      Left            =   7440
      Picture         =   "FormHint.frx":3C5E
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1470
   End
End
Attribute VB_Name = "FormHint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub bne_Click()
FormHint.Visible = False
FormMainMode.wmpse3.Controls.stop
FormMainMode.wmpse3.Controls.play
一般系統類.檢查音樂播放 3
選單使用者事件 = True
選單電腦事件 = True
End Sub

Private Sub bnet_Click()
FormHint.Visible = False
FormMainMode.wmpse3.Controls.stop
FormMainMode.wmpse3.Controls.play
一般系統類.檢查音樂播放 3
選單使用者事件 = True
選單電腦事件 = True
End Sub

