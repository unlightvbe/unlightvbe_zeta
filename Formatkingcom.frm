VERSION 5.00
Begin VB.Form Formatkingcom 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "UnlightVBE-技能啟動中"
   ClientHeight    =   9945
   ClientLeft      =   9480
   ClientTop       =   1965
   ClientWidth     =   6780
   Icon            =   "Formatkingcom.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   6780
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   8040
   End
   Begin VB.PictureBox atkingcomjpg 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   11025
      Left            =   -840
      Picture         =   "Formatkingcom.frx":0CCA
      ScaleHeight     =   11025
      ScaleWidth      =   18600
      TabIndex        =   0
      Top             =   -600
      Width           =   18600
   End
End
Attribute VB_Name = "Formatkingcom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim d As Integer

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
一般系統類.離開遊戲提示 Cancel, UnloadMode
End Sub

Private Sub t1_Timer()
If 目前數(31) = 19 Then
   Formatkingcom.Visible = False
   t1.Enabled = False
   If Val(FormMainMode.atkingnumtot.Caption) > 0 Then
      If atkingno(Val(FormMainMode.atkingnumtot.Caption), 11) = 0 Then
          If Formsetting.checktest.Value = 1 Then Debug.Print "Formatkingcom If atkingno(,11) = 0 後"
          '=======================
          FormMainMode.atkingnumtot.Caption = Val(FormMainMode.atkingnumtot.Caption) - 1
          FormMainMode.atkingtrtot.Interval = 20
          FormMainMode.atkingtrtot.Enabled = True
      End If
   End If
   Unload Me
ElseIf 目前數(31) = 10 Then
   FormMainMode.技能執行中更換圖片_Timer
   FormMainMode.技能執行中啟動_Timer
   目前數(31) = Val(目前數(31)) + 1
ElseIf 目前數(31) = 7 Then
   FormMainMode.wmpse5.Controls.play
   一般系統類.檢查音樂播放 5
   目前數(31) = Val(目前數(31)) + 1
ElseIf 目前數(31) = 5 Then
   atkingcomjpg.Visible = True
   目前數(31) = Val(目前數(31)) + 1
Else
   目前數(31) = Val(目前數(31)) + 1
End If
End Sub
